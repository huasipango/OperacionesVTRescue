VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFacturasComisiones 
   Caption         =   "Facturación de Comisiones Bono Electrónico"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   10680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   550
      Left            =   120
      TabIndex        =   17
      Top             =   1320
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
         ItemData        =   "frmFacturasComisiones.frx":0000
         Left            =   1800
         List            =   "frmFacturasComisiones.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   120
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
         TabIndex        =   19
         Top             =   180
         Width           =   1545
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   255
      Left            =   8520
      TabIndex        =   16
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdActualizaFolio 
      Caption         =   "Actualiza Folio"
      Height          =   375
      Left            =   9000
      TabIndex        =   15
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   10335
      Begin FPSpread.vaSpread spdFacturas 
         Height          =   3735
         Left            =   120
         OleObjectBlob   =   "frmFacturasComisiones.frx":001C
         TabIndex        =   6
         Top             =   240
         Width           =   10095
      End
   End
   Begin VB.Frame fraPeriodo1 
      Caption         =   "Periodo a Facturar  (dd/mm/yyyy) "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   10335
      Begin VB.CheckBox chkSelTodas 
         Alignment       =   1  'Right Justify
         Caption         =   "Seleccionar Todas"
         Height          =   255
         Left            =   8400
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   8400
         Picture         =   "frmFacturasComisiones.frx":1251
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   450
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   9600
         Picture         =   "frmFacturasComisiones.frx":1353
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   9000
         Picture         =   "frmFacturasComisiones.frx":1455
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1200
         TabIndex        =   1
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   850
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
         TabIndex        =   2
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   855
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
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   855
         Width           =   255
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   850
         Width           =   495
      End
   End
   Begin VB.Label Label3 
      Caption         =   " * Doble click sobre el grupo que desea para ver sus transacciones detalladas"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Width           =   5535
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      X1              =   3120
      X2              =   10440
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   240
      Picture         =   "frmFacturasComisiones.frx":1557
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Facturación de Comisiones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmFacturasComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated

Private Sub cboBodegas_Click()
  Call STAT
End Sub

Private Sub cmdAbrir_Click()
 With spdFacturas
   Screen.MousePointer = 11
   .Col = -1
   .Row = -1
   .Action = 12
   .MaxRows = 0
         
   sqls = "sp_Consultas_CargosBanco '" & Format(mskFechaIni, "MM/DD/YYYY") & "','" & Format(mskFechaFin, "MM/DD/YYYY") & "',"
   sqls = sqls & Product & ",'ComisionesAbre'," & cboBodegas.ItemData(cboBodegas.ListIndex)
         
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
   
   If rsBD.EOF Then
      MsgBox "No hay transacciones en este rango de fechas", vbCritical, "No hay transacciones"
      Screen.MousePointer = 1
      Exit Sub
   End If
   
   i = 1
   Do While Not rsBD.EOF
      .MaxRows = i
      .Row = i
      .Col = 1
      .Text = rsBD!Grupo
      .Col = 2
      .Text = rsBD!descripcion
      .Col = 3
      .Text = rsBD!totaltran
      .Col = 4
      .Text = Format(rsBD!Comision, "##.00")
      .Col = 5
      .Text = rsBD!TotalComision
      .Col = 6
      .Text = rsBD!Ivacom
      .Col = 7
      .Text = rsBD!totalfactura
      i = i + 1
      rsBD.MoveNext
   Loop
   
   rsBD.Close
   Set rsBD = Nothing
   Screen.MousePointer = 1
 End With
End Sub

Private Sub cmdActualizaFolio_Click()
Dim resp As Integer
Dim Folio As Long

   sqls = "sp_FoliosBE " & cboBodegas.ItemData(cboBodegas.ListIndex) & ",Null,'Busca'"
   
'  sqls = " select Prefijo serie, consecutivo " & _
'         " From folios " & _
'         " Where Bodega =" & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'         " and tipo = 'FCM'"
   
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic
   
   If Not rsBD.EOF Then
      Folio = rsBD!consecutivo
   Else
      MsgBox "Error en folio de la Factura, verifique con sistemas", vbCritical
      Exit Sub
   End If
      
On Error GoTo err_panel
   resp = InputBox("Folio de Factura Nuevo", "Folio de Facturas de Comisiones", Folio)
   
On Error GoTo err_gral

   sqls = "sp_FoliosBE " & cboBodegas.ItemData(cboBodegas.ListIndex) & "," & resp & ",'Actualiza'"
   
'   sqls = "update folios set consecutivo = " & RESP & _
'          " , fechamodificacion = getdate()" & _
'          " Where Bodega =" & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'          " and tipo = 'FCM'"
   
   cnxBD.Execute sqls, intRegistros
   
   Exit Sub
   
err_panel:
         Exit Sub
err_gral:
         MsgBox ERR.Description
         Exit Sub
End Sub

Private Sub cmdGrabar_Click()
   GrabaFactura
End Sub
Function Imprimefactura(Bodega As Integer, Grupo As Integer, Factura As Long, Serie As String)
Dim archfisico As String
Dim ArchImpre
Dim fechaini, fechafin, Fecha As String
Dim Nombre, Domicilio, Colonia, Rfc, poblacion, estado As String
Dim Comision, Ivacom, TotalF
Dim svalor As String
Dim sEdo As String, spobl As String, scp As String

On Error GoTo err_gral

sqls = " select a.factura, b.descripcion Nombre, b.domicilio, b.colonia," & _
       " b.rfc,b.telefono,c.descripcion Poblacion, d.desccorta estado," & _
       " a.serie, a.cliente, a.fecha  ,a.fechaini, a.fechafin," & _
       " a.subtotal + a.iva ImporteF,a.subtotal comision, a.iva ivacomision" & _
       " from fm_facturas a, grupos b, poblaciones c, estados d" & _
       " Where a.Bodega = " & Bodega & _
       " and a.cliente = " & Grupo & _
       " and a.serie = '" & Serie & "'" & _
       " and a.factura =  " & Factura & _
       " and a.cliente = b.cveestablecimiento" & _
       " AND b.Producto=" & Product & _
       " and b.poblacion = c.poblacion" & _
       " and b.estado = d.estado"

 Set rsBD = New ADODB.Recordset
 rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
 
 If Not rsBD.EOF Then
   Close #2
   
   ArchImpre = "c:\Facturacion\f" & Factura & ".txt"
   Open ArchImpre For Output As #2
                  
   svalor = Leyenda(rsBD!ImporteF)
   
   GoSub Encabezado
   GoSub Detalle
   GoSub PieDePagina
 End If

   Exit Function
   
Encabezado:

   sEdo = Trim(rsBD!estado & "")
   spobl = Trim(rsBD!poblacion & "")
   scp = 0
   
   spobl = Left(spobl, 40 - (Len(sEdo) + 2) - (Len(scp) + 2)) & ", " & sEdo & "  " & scp

   For i = 1 To 11
      Print #2,
   Next
   Print #2, Tab(5); Pad(Trim(Mid(rsBD!Nombre, 1, 30)), 30, " ", "R"); _
               Tab(50); Pad(Trim(IIf(IsNull(rsBD!Rfc), " ", rsBD!Rfc)), 15, " ", "R"); _
               Tab(75); Pad(Trim(Format(rsBD!cliente, "#####")), 5, " ", "R"); _
               Tab(86); Pad(Trim(Format(rsBD!Fecha, "YYYY-MM-DD")), 10, " ", "R"); _
               Tab(100); Pad(Trim(rsBD!Serie), 2, " ", "R") & Pad(Trim(Format(rsBD!Factura, "#######")), 7, " ", "R")
   Print #2, Tab(5); Pad(Trim(Mid(rsBD!Nombre, 31, 30)), 30, " ", "R")
   Print #2, Tab(5); rsBD!Domicilio
   Print #2, Tab(5); rsBD!Colonia
   Print #2, Tab(5); Pad(Trim(spobl), 40, " ", "R")
   Print #2, Tab(50); Pad(Trim(rsBD!Telefono), 14, " ", "R"); _
             Tab(75); Pad(Trim(Format(0, "########")), 8, " ", "R");
               'Tab(105); "R" & rsbd1!RutaNumero
   For i = 1 To 5
      Print #2,
   Next
   Print #2, ' Tab(5); "----------------------- DESGLOSE DEL PEDIDO ----------------------- "
   Print #2,

   Return
   
Detalle:

    Print #2, Tab(7); "1   COBRO DE COMISION DE LAS TRANSACCIONES";
    Print #2, Tab(7); "    REALIZADAS DEL " & UCase(Format(rsBD!fechaini, "DD/MMM/YY")) & " AL " & UCase(Format(rsBD!fechafin, "DD/MMM/YY")) & ""; _
             Tab(60); "1"; Tab(76); Format(rsBD!Comision, "###,##0.00"); Tab(89); Format(rsBD!ivacomision, "#,##0.00"); _
             Tab(100); Format(rsBD!ImporteF, "###,##0.00")
   

    For i = 3 To 20 - cont
        Print #2,
    Next i
    
   Return
PieDePagina:

    sqls = "select impuestointerior iva from bodegas where bodega = " & Bodega
    
    Set rsiva = New ADODB.Recordset
    rsiva.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsiva.EOF Then
        iva = rsiva!iva
    Else
        iva = 0
    End If
    
   Print #2, 'Tab(81); Format(rsBD!Comision, "###,##0.00"); Tab(90); Format(rsBD!ivacomision, "#,##0.00"); Tab(101); Format(rsBD!ImporteF, "###,##0.00")
 
   Print #2, Tab(10); "IVA: " & iva & " %";
   Print #2, ' Tab(20); "SUBTOTAL : "; _
                     Tab(72); Pad(Trim(Format(Subtotal, "#,###,##0.00")), 12, " ", "L"); _
                     Tab(85); Pad(Trim(Format(ivaexe, "##,##0.00")), 9, " ", "L"); _
                     Tab(96); Pad(Trim(Format(total1, "##,###,##0.00")), 13, " ", "L")
   Print #2, 'Tab(20); sServicio; _
                     Tab(72); Pad(Trim(Format(val_comision, "###,##0.00")), 12, " ", "L"); _
                     Tab(85); Pad(Trim(Format(ivacomis, "##,##0.00")), 9, " ", "L"); _
                     Tab(96); Pad(Trim(Format(total2, "##,###,##0.00")), 13, " ", "L")
   Print #2,
   Print #2,
   Print #2,
   
   Print #2, Tab(60); "1"; Tab(76); Format(rsBD!Comision, "###,##0.00"); _
                     Tab(89); Format(rsBD!ivacomision, "#,##0.00"); _
                     Tab(100); Format(rsBD!ImporteF, "###,##0.00")
   For lngI = 1 To 5
      Print #2,
   Next

   Print #2, Tab(17); "PAGO EN UNA SOLA EXHIBICION";
   
   svalor = Pad(Trim(Format(rsBD!ImporteF, "$#,###,##0.00")), 12, " ", "L") & " " & svalor
   
   Print #2, Tab(17); "www.valetotal.com"
   Print #2, Tab(70); Mid(Trim(svalor), 1, 40)
   If Len(svalor) >= 40 Then
       Print #2, Tab(70); Mid(Trim(svalor), 41, Len(Trim(svalor)) - 40);
   End If
   
   Print #2,

   Print #2,
   
   Close #2
   
   Return

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmFacturasComisiones.ImprimeFactura")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Impresion de Facturas de Comisiones"
  'If cnxBD.BeginTrans > 0 Then cnxBD.RollbackTrans
'  Resume Next
   MsgBar "", False
    
End Function

Sub GrabaFactura()

Dim resp
Dim impuesto As Double
Dim sserie As String
Dim Folio As Long, Grupo As Integer
Dim ImporteT, ImporteF, Comision, Ivacom, TotComision As Double
Dim resultado
Dim reg As Integer
Dim cliente As Integer

On Error GoTo err_gral
      
PRIMERA = True

With spdFacturas

 For reg = 1 To .MaxRows
 
     Screen.MousePointer = 11

    .Row = reg
    .Col = 8
    If .value = 1 Then
      .Col = 1
      Grupo = .Text
      .Col = 3
      ImporteT = CDbl(.Text)
      .Col = 4
      Comision = CDbl(.Text)
      .Col = 5
      TotComision = CDbl(.Text)
      .Col = 6
      Ivacom = CDbl(.Text)
      .Col = 7
      ImporteF = CDbl(.Text)
      
      'cnxBD.BeginTrans
      
      sqls = "sp_FoliosBE " & cboBodegas.ItemData(cboBodegas.ListIndex) & ",Null,'Busca'"
      
'      sqls = " select Prefijo serie, consecutivo " & _
'            " From folios " & _
'            " Where Bodega =" & cboBodegas.ItemData(cboBodegas.ListIndex) & _
'            " and tipo = 'FCM'"
            
      Set rsBD = New ADODB.Recordset
      rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic
    
      If rsBD.EOF = False Then
         sserie = Trim(rsBD!Serie)
         Folio = rsBD!consecutivo + 1
      Else
         cnxBD.RollbackTrans
         MsgBar "", False
         MsgBox "Falta valor de la serie ", vbOKOnly, "Avisar a Sistemas"
         Exit Sub
      End If
              
      If sserie = "" Then
   '      cnxBD.RollbackTrans
         MsgBar "", False
         MsgBox "Falta valor de la serie, Verifique con Sistemas  ", vbOKOnly, "Avisar a Sistemas"
         Exit Sub
      End If
           
        
      sqls = " select impuestointerior" & _
             " From bodegas " & _
             " Where Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
             
      Set rsimp = New ADODB.Recordset
      rsimp.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic
            
      If Not rsBD.EOF Then
         impuesto = (rsimp!ImpuestoInterior) / 100
      Else
         impuesto = ivagral '0.15
      End If
      
     
      '----  guarda en la misma tabla de comisiones papel
      
        sqls = " select isnull(cveestablecimiento,0) cveestab from grupos" & _
               " where grupo = " & Grupo
        sqls = sqls & " and Producto=" & Product
               
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
        
        
        If Not rsBD.EOF Then
            cliente = Val(rsBD!cveestab)
                   
        Else
            MsgBox "No se encontró el cliente a facturar", vbInformation, "Cliente no encontrado"
    '         cnxBD.RollbackTrans
             Exit Sub
        End If
        
        If cliente = 0 Then
            MsgBox "El Grupo " & Grupo & " no tiene dado de alta su clave de establecimiento, no se puede facturar", vbCritical, "Grupo no dado de alta"
     '       cnxBD.RollbackTrans
        Else
               
            sqls = " EXEC sp_fm_Clientes_mov_ins "
            sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
            sqls = sqls & vbCr & ", @Cliente      = " & Val(cliente)
            sqls = sqls & vbCr & ", @Fecha        = '" & Format(Date, "MM/DD/YYYY") & "'"
            sqls = sqls & vbCr & ", @Tipo_Mov     = 21" 'Facturas comisiones be
            sqls = sqls & vbCr & ", @Serie        = '" & Trim(sserie) & "'"
            sqls = sqls & vbCr & ", @Refer        = 2"
            sqls = sqls & vbCr & ", @Refer_Apl    = " & Folio
            sqls = sqls & vbCr & ", @CarAbo       = 'C'"
            sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
            
            sqls = sqls & vbCr & ", @Importe      = " & TotComision
            sqls = sqls & vbCr & ", @iva      = " & Ivacom
            sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
            
            sqls = sqls & vbCr & ", @Fecha_Mov =  '" & Format(Date, "MM/DD/YYYY") & "'"
            sqls = sqls & vbCr & ", @Usuario = '" & gstrUsuario & "'"
            sqls = sqls & vbCr & ", @TipoBon=" & Product
            cnxBD.Execute sqls, intRegistros
            
            sqls = " exec sp_FM_facturas @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                       " ,@AnoFactura = " & Year(Date) & _
                       " ,@Serie = '" & Trim(sserie) & "'" & _
                       " ,@Factura = " & Folio & _
                       " ,@Cliente = " & Val(cliente) & _
                       " ,@Fecha    =  '" & Format(Date, "mm/dd/yyyy") & "' " & _
                       " ,@Subtotal = " & TotComision & _
                       " ,@Iva    = " & Ivacom & _
                       " ,@Rubro = 21 " & _
                       " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                       " ,@Status = 1" & _
                       " ,@StatusImpreso = 0" & _
                       " ,@Fechaini        =    '" & Format(mskFechaIni, "mm/dd/yyyy") & "'" & _
                       " ,@Fechafin        =    '" & Format(mskFechaFin, "mm/dd/yyyy") & "'"
     
            cnxBD.Execute sqls, intRegistros
            
            'Actualiza los status de transacciones ya facturadas
            sqls = " EXEC sp_act_tran_x_comdet"
            sqls = sqls & vbCr & " @Grupo      = " & Grupo
            sqls = sqls & vbCr & ", @Fechaini        =    '" & Format(mskFechaIni, "mm/dd/yyyy") & "'"
            sqls = sqls & vbCr & ", @Fechafin        =    '" & Format(mskFechaFin, "mm/dd/yyyy") & "'"
            sqls = sqls & vbCr & ", @Producto        =    " & Product
            
            cnxBD.Execute sqls, intRegistros
            
            
          sqls = " update folios set consecutivo   = " & Folio & _
                " Where Bodega =" & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                " and tipo = 'FCM'"
    
          cnxBD.Execute sqls, intRegistros
      '    cnxBD.CommitTrans
          
          blnreimp = False
          
'          resultado = Imprimefactura(cboBodegas.ItemData(cboBodegas.ListIndex), Cliente, folio, sserie)
          Call doGenArchFE_OI(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(sserie), Folio, Folio)
         Screen.MousePointer = 1

        End If
    End If
Next reg
End With
ImpUnaVez = False

cmdAbrir_Click

Exit Sub

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmFacturaComisiones.GrabaFactura")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Asignación de Tarjetas"
  'If cnxBD.BeginTrans > 0 Then cnxBD.RollbackTrans
  ' Resume Next
   MsgBar "", False
End Sub
Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Command1_Click()
   Call doGenArchFE_OI(1, "VA", 12, 12)
End Sub

Private Sub chkSelTodas_Click()
With spdFacturas
If chkSelTodas.value = 1 Then
        For i = 1 To .MaxRows
            .Row = i
            .Col = 8
            .value = 1
        Next i
    Else
        For i = 1 To .MaxRows
            .Row = i
            .Col = 8
            .value = 0
        Next i
    End If
End With
End Sub

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
  'Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Sub InicializaForma()
   spdFacturas.MaxRows = 1
   Call CargaBodegas(cboBodegas)
End Sub
Private Sub Form_Activate()
 Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & Trim(CboProducto.Text) & " ','Leer'")
End Sub
Private Sub Form_Load()
   Set mclsAniform = New clsAnimated
'   Set cnxBD = New ADODB.Connection
'   cnxBD.CommandTimeout = 2000
'   cnxBD.Open "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
   mskFechaIni.Text = Date
   mskFechaFin.Text = Date
   CboProducto.Clear
   Call CargaComboBE(CboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(CboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(CboProducto.Text)) & " ','Leer'")
   CboProducto.Text = UCase("Winko Mart")
   Call CargaBodegas(cboBodegas)
   ImpUnaVez = False
   Call STAT
End Sub
Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptTrans_x_com.rpt"
    mdiMain.cryReport.Destination = Destino
    spdFacturas.Row = spdFacturas.ActiveRow
    spdFacturas.Col = 1
    mdiMain.cryReport.StoredProcParam(0) = Val(spdFacturas.Text)
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "mm/dd/yyyy")
    mdiMain.cryReport.StoredProcParam(3) = CboProducto.ListIndex + 1
    mdiMain.cryReport.StoredProcParam(4) = 0
        
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'  cnxBD.Close
'  Set cnxBD = Nothing
  Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub

Private Sub spdFacturas_DblClick(ByVal Col As Long, ByVal Row As Long)

With spdFacturas
    If Row = 0 Then
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
    Else
   
    If Col = 1 Then
      Imprime crptToWindow
    End If
   End If
   
End With
End Sub

Sub STAT()
    cmdGrabar.Enabled = True
End Sub
