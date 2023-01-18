VERSION 5.00
Begin VB.Form frmNotasCargo 
   Caption         =   "Vale Total, S.A. de C.V."
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frBotones 
      Height          =   855
      Left            =   240
      TabIndex        =   18
      Top             =   4320
      Width           =   9015
      Begin VB.CommandButton Command1 
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdActualizaFolio 
         Caption         =   "Actualiza Folio"
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
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   8160
         Picture         =   "frmNotasCargo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   525
         Left            =   7320
         Picture         =   "frmNotasCargo.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frCte 
      Height          =   3375
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   9015
      Begin VB.TextBox txtImporte 
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtIva 
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
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtTotal 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox txtFolio 
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
         Left            =   7320
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtSerie 
         Alignment       =   2  'Center
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
         Left            =   3720
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1200
         Width           =   500
      End
      Begin VB.TextBox txtFactura 
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
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   7320
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   375
         Left            =   2400
         Picture         =   "frmNotasCargo.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   720
         Width           =   375
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
         Left            =   1200
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmNotasCargo.frx":0306
         Left            =   1200
         List            =   "frmNotasCargo.frx":030D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label16 
         Caption         =   "Importe:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "IVA:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "T O T A L :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label lblPlaza 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lblNombre 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   21
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "Sig. Folio:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Serie:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Factura:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   15
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   9120
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Notas de Cargo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "frmNotasCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumeroNota As Integer
Public iva As Integer

Private Sub cboBodegas_Click()
    sqls = " select serie_factura as serie, impuestointerior as iva from bodegas"
    sqls = sqls & " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
    Set rsbod = New ADODB.Recordset
    rsbod.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic
    
    If Not rsbod.EOF Then
        txtserie = rsbod!serie
        iva = rsbod!iva '15
    Else
        MsgBox "Error en la serie de la sucursal"
    End If
    
    rsbod.Close
    Set rsbod = Nothing
    
End Sub

Private Sub cmdActualizaFolio_Click()
Dim resp As Integer
Dim Folio As Integer

   sqls = " select Prefijo serie, consecutivo " & _
         " From folios " & _
         " Where Bodega =0 and tipo = 'NCA'"
   Set rsBD = New ADODB.Recordset
   rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockPessimistic
   
   If Not rsBD.EOF Then
      Folio = rsBD!consecutivo
   Else
      MsgBox "Error en folio de las Notas de Cargo, verifique con sistemas", vbCritical
      Exit Sub
   End If
      
On Error GoTo err_panel
   resp = InputBox("Folio de Nota de Cargo Nuevo", "Folio de Notas de Cargo", Folio)
   
On Error GoTo err_gral

   sqls = "update folios set consecutivo = " & resp & _
          " , fechamodificacion = getdate()" & _
          " Where Bodega = 0 and tipo = 'NCA'"
   
   cnxBD.Execute sqls, intRegistros
   
   txtFolio = resp + 1
   
   Exit Sub
   
err_panel:
         Exit Sub
err_gral:
         MsgBox ERR.Description
         Exit Sub
   
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
    
    TipoBusqueda = "Cliente"
    frmConsulta.Show vbModal
    
    If frmConsulta.cliente >= 0 Then
       txtCliente = frmConsulta.cliente
       lblNombre = frmConsulta.Nombre
    End If
    Set frmConsulta = Nothing
    MsgBar "", False
End Sub
Private Sub cmdGrabar_Click()

Dim TIPO_MOV, Concepto As Integer
Dim transO As Boolean
    
On Error GoTo err_gral

   
    NumeroNota = 1
    sqls = " select * from folios"
    sqls = sqls & " where bodega   = 0" 'Laredo lleva el mismo folio que monterrey
    sqls = sqls & " and tipo = 'NCA'"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        serie = Trim(rsBD!Prefijo)
        txtFolio = Val(rsBD!consecutivo) + 1
        TIPO_MOV = 43
        
        sqls = " EXEC sp_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & txtCliente
        sqls = sqls & vbCr & ", @Fecha        =    '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = " & TIPO_MOV
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(txtserie) & "'"  'Serie Nota Credito
        sqls = sqls & vbCr & ", @Refer        = " & txtFolio
        sqls = sqls & vbCr & ", @Cuenta_origen = " & txtFolio
        sqls = sqls & vbCr & ", @Refer_Apl    = " & txtFactura
        sqls = sqls & vbCr & ", @CarAbo       = 'C'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & txtTotal
        sqls = sqls & vbCr & ", @Ficha     = " & txtFolio
        sqls = sqls & vbCr & ", @CreditoFac = 'N'"
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = " & Usuario
               
        cnxBD.Execute sqls, intRegistros
        
        
        sqls = "exec sp_NotasCar "
        sqls = sqls & vbCr & "   @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & "  ,@Folio        = " & txtFolio
        sqls = sqls & vbCr & "  ,@Serie        = '" & Trim(txtserie) & "'"
        sqls = sqls & vbCr & "  ,@cliente      = " & txtCliente & ""
        sqls = sqls & vbCr & "  ,@Factura      = " & txtFactura
        sqls = sqls & vbCr & "  ,@TipoMov      = " & TIPO_MOV
        sqls = sqls & vbCr & "  ,@Concepto      = " & 10   'concepto nota cargo
        sqls = sqls & vbCr & "  ,@FechaEmi      = '" & txtFecha & "'"
        sqls = sqls & vbCr & "  ,@TpoBon      =  0"
        sqls = sqls & vbCr & "  ,@Bonexe      = " & Val(txtImporte)
        sqls = sqls & vbCr & "  ,@Bongra     = 0"
        sqls = sqls & vbCr & "  ,@Ivagra      = 0"
        sqls = sqls & vbCr & "  ,@Comision      = 0"
        sqls = sqls & vbCr & "  ,@IvaCom      = " & txtIva
        sqls = sqls & vbCr & "  ,@Valor      = " & txtTotal
        sqls = sqls & vbCr & "  ,@status      = 0"
        sqls = sqls & vbCr & "  ,@Reembolso      =0"
        sqls = sqls & vbCr & "  ,@Usuario      = '" & Usuario & "'"
        sqls = sqls & vbCr & "  ,@Plaza      = ' '"
        sqls = sqls & vbCr & "  ,@TipoBon      = 0"
        sqls = sqls & vbCr & "  ,@SerieNota      = '" & Trim(serie) & "'"

                
        cnxBD.Execute sqls, intRegistros
        
        sqls = "update folios set consecutivo = " & Val(txtFolio) & _
               " where bodega = 0 and tipo = 'NCA'"
               
         cnxBD.Execute sqls, intRegistros
         
         
    Else
        MsgBox "Error al buscar el folio, verifique con Sistemas"
        cnxBD.RollbackTrans
        Exit Sub
    End If
    
    rsBD.Close
    Set rsBD = Nothing
      
   
    'Call ImprimeNota(cboBodegas.ItemData(cboBodegas.ListIndex), Val(txtFolio.Text))
    Call doGenArchFE_NCA(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(serie), Val(txtFolio.Text), Val(txtFolio.Text), 4)
       
    NumeroNota = NumeroNota + 1
  
    LimpiarControles Me
        
    txtFecha = Format(Date, "MM/DD/YYYY")
    CargaBodegas cboBodegas
    BuscaSigFolio (cboBodegas.ItemData(cboBodegas.ListIndex))
    Exit Sub

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmNotasCargo.cmdGrabar")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Notas de Credito"
'   If transO = True Then cnxBD.RollbackTrans
   MsgBar "", False
        
End Sub
Sub ImprimeNota(Bodega As Integer, Folio As Integer)
    
    sqls = "select a.folio, a.serie, a.cliente, b.nombre, b.RFC, b.domicilio,"
    sqls = sqls & " c.descripcion AS POBLA ,a.factura, a.valor, a.FechaEmi,"
    sqls = sqls & " a.BonExe , a.BonGra, a.IvaGra, a.Comision, a.IvaCom, a.plaza"
    sqls = sqls & " from   notascar a, clientes b, poblaciones c"
    sqls = sqls & " Where a.Cliente = b.Cliente"
    sqls = sqls & " and b.poblacion = c.poblacion"
    sqls = sqls & " and a.bodega = " & Bodega
    sqls = sqls & " and a.folio = " & Folio
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    If Not rsBD.EOF Then
        
        sql = " select fecha from clientes_movimientos"
        sql = sql & " Where Bodega = " & Bodega
        sql = sql & " and refer_Apl = " & rsBD!Factura
        sql = sql & " and cliente = " & rsBD!cliente
        sql = sql & " and tipo_mov = 10"
    
        Set rsBD2 = New ADODB.Recordset
        rsBD2.Open sql, cnxBD, adOpenForwardOnly, adLockReadOnly
    
        If Not rsBD2.EOF Then
            FechaFac = rsBD2!Fecha
        Else
            Fecha = ""
        End If
    
        If NumeroNota = 1 Then
           Print #2,
        End If
        
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        Print #2,
        
        Print #2, Tab(29); "NOTA DE CARGO     FOLIO: " & Format(rsBD!serie, "!@@") & " " & Format(rsBD!Folio, "!@@@@@@")
        Print #2,
        Print #2, Tab(17); "CIA.: " & Format(rsBD!cliente, "@@@@@@")
        Print #2, Tab(17); "SR.(ES): " & RTrim(Mid(rsBD!Nombre, 1, 30)) + " " + LTrim(Mid(rsBD!Nombre, 31, 60))
        Print #2, Tab(26); Trim(rsBD!Plaza)
        Print #2, Tab(17); "R.F.C.: " & Trim(rsBD!Rfc); Tab(51); "VALOR: $ " & Format(Format(rsBD!valor, "###,###,###,###.00"), "@@@@@@@@@@@@@@@")
        Print #2, Tab(17); Left(rsBD!Domicilio, 33); Tab(51); "FECHA: " & Format(rsBD!fechaemi, "YYYY-MM-DD")
        Print #2, Tab(17); Left(rsBD!POBLA, 60)
        Print #2, Tab(17); "CON ESTA FECHA HEMOS CARGADO A SU CUENTA LO SIGUIENTE:"
        Print #2, Tab(17); "| ----------------------------------------------------------------- |"
        Print #2, Tab(17); "|            DISTRIBUCION          |             CONCEPTO           |"
        Print #2, Tab(17); "| ----------------------------------------------------------------- |"
        
             
        Print #2, Tab(17); "| IMPORTE      :  $" & Format(Format(IIf(IsNull(rsBD!BONEXE), 0, rsBD!BONEXE), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| NOTA DE CARGO POR CHEQUE DEV.  |"
        Print #2, Tab(17); "| IVA          :  $" & Format(Format(IIf(IsNull(rsBD!Ivacom), 0, rsBD!Ivacom), "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "|                                |"
        Print #2, Tab(17); "|                                  "; Tab(52); "| FACTURA:" & Format(rsBD!Factura, "@@@@@@"); Tab(74); Format(FechaFac, "YYYY-MM-DD"); " |"
        Print #2, Tab(17); "|                                  |                                |"
        
        
        
        
        Leyenda (rsBD!valor), ""
        letrero = "(" & letrero & ")"
        texto1 = Mid(letrero, 1, 30)
        texto2 = Mid(letrero, 31, 30)
        texto3 = Mid(letrero, 62, 30)
        
        
        Print #2, Tab(17); "|                                  "; Tab(52); "| " & Format(texto1, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "| T O T A L:      $" & Format(Format(rsBD!valor, "###,###,###,##0.00"), "@@@@@@@@@@@@@@@"); Tab(52); "| " & Format(IIf(texto2 = "", " ", texto2), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "|                                  "; Tab(52); "| " & Format(IIf(texto3 = "", " ", texto3), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") & " |"
        Print #2, Tab(17); "|                                  |                                |"
        Print #2, Tab(17); "| -------------------------------- | ------------------------------ |"
        Print #2,

        
   End If
   NumeroNota = NumeroNota + 1
        
    
End Sub
Private Sub cmdSalir_Click()

   ' Close #2
'
'    resp = MsgBox("Desea imprimir las Notas de Cargo capturadas?", vbYesNo, "Notas de Crédito")
'
'    If resp = vbYes Then
'        Close #14
'        archfisico = "C:\facturacion\NotasCar.bat"
'        Open archfisico For Output As #14
'
'        SQL = " SELECT impresora FROM BON_IMPRESORAs WHERE doCumento = 'FACTURAS' AND MAQUINA = " & gnMaquina
'        Set rsbd = New ADODB.Recordset
'        rsbd.Open SQL, cnxBD, adOpenForwardOnly, adLockReadOnly
'        If Not rsbd.EOF Then
'            Print #14, "COPY C:\facturacion\notascar.lis " & Trim(rsbd!impresora)
'        Else
'            Print #14, "copy C:\facturacion\notascar.lis lpt1"
'        End If
'        rsbd.Close
'
'        Close #14
'        Shell "C:\facturacion\NotasCar.BAT", 2
'    End If
    
    
        '-------------------------
    
'  sTipoArch = "NOTAS"
'  strPuerto = doFindPrinter(gstrPC, sTipoArch)
'  nFileBat = FreeFile()
'  sFileBat = "C:\facturacion\Notascar.bat"
'  SFileFact = "C:\Facturacion\notascar.lis"
'
'
'
'  Open sFileBat For Output As #nFileBat
'  If gsOS = "XP" Then
'     Print #nFileBat, "PRINT /D:" & strPuerto & " " & SFileFact
'  Else
'     Print #nFileBat, "COPY " & SFileFact & " " & strPuerto
'  End If
'  Close #nFileBat
'  If MsgBox("¿Desea mandar las notas de cargo a la impresora?", vbYesNo) = vbYes Then
'      EsperarShell (sFileBat) '--, vbMinimizedFocus
'  End If
'  DoEvents
'
'  Kill SFileFact
'
'
'
'
    
    '-------------------------
    
    Unload Me
End Sub

Private Sub Command1_Click()
'Call doGenArchFE_NCA(13, "NC", 4, 4, 4)
'Call doGenArchFE_NCA(1, "NC", 5, 5, 4)
'Call doGenArchFE_NCA(2, "NC", 3, 3, 4)


End Sub

Private Sub Form_Load()

  
    Limpia_Campos
    txtFecha = Format(Date, "MM/DD/YYYY")
   
    CargaBodegas cboBodegas
       
    
'   Abre el archivo para generar las notasde crédito facturada
    Dim ArchImpre
    Close #2
    ArchImpre = "C:\Facturacion\notascar.lis"
    Open ArchImpre For Output As #2
    NumeroNota = 1
    
    BuscaSigFolio (cboBodegas.ItemData(cboBodegas.ListIndex))
    
    
End Sub

Sub Limpia_Campos()
    LimpiarControles Me
End Sub
Sub BuscaSigFolio(Bodega As Integer)

    
    sqls = " select isnull(consecutivo,0) consecutivo from folios"
    sqls = sqls & " where bodega   =0 "
    sqls = sqls & " and tipo = 'NCA'"
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rsBD.EOF Then
      txtFolio = Val(rsBD!consecutivo) + 1
    Else
      txtFolio = 0
      MsgBox "Error en el folio", vbCritical, "Notas de Cargo"
      
    End If

        
End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCliente_LostFocus()
    sqls = "select nombre from clientes where cliente = " & Val(txtCliente)
    
    Set rsBD = New ADODB.Recordset
    rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
    
    
    If Not rsBD.EOF Then
        lblNombre = rsBD!Nombre
    End If
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtFactura <> "" Then
            If txtCliente = "" Then
                MsgBox "Primero debe capturar el cliente!", vbInformation
                txtFactura = ""
                txtCliente.SetFocus
                Exit Sub
            End If
            sqls = " select * from clientes_movimientos" & _
                   " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
                   " and cliente = " & Val(txtCliente) & _
                   " and refer_Apl = " & Val(txtFactura)
                   
            Set rsBD = New ADODB.Recordset
            rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly
            
            If rsBD.EOF Then
                MsgBox "La factura no existe o no corresponde a este cliente", vbCritical, "Notas de Cargo"
                txtFactura = ""
                txtFactura.SetFocus
            Else
                SendKeys "{TAB}"
            
            End If
            
        End If
    End If
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtImporte <> "" Then
            txtIva = Val(txtImporte) * (iva / 100)
            txtTotal = Val(txtImporte) + Val(txtIva)
            SendKeys "{tab}"
        End If
    End If
End Sub

