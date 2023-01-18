VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#2.5#0"; "Ss32x25.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOI 
   Caption         =   "Facturacion"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtobs 
      Height          =   285
      Left            =   240
      MaxLength       =   50
      TabIndex        =   55
      Top             =   6960
      Width           =   6495
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza Folio"
      Height          =   375
      Left            =   7320
      TabIndex        =   41
      Top             =   6720
      Width           =   1455
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
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8535
      Begin VB.ComboBox cboProductos 
         Height          =   315
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1080
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   7320
         TabIndex        =   51
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   6000
         Picture         =   "frmOI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Eliminar"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   5400
         Picture         =   "frmOI.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   450
      End
      Begin VB.TextBox txtCliente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarC 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   350
         Left            =   2160
         Picture         =   "frmOI.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1500
         Width           =   350
      End
      Begin VB.ComboBox cboConcep 
         Height          =   315
         ItemData        =   "frmOI.frx":0306
         Left            =   1080
         List            =   "frmOI.frx":0308
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton cmdGrabar 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   7200
         Picture         =   "frmOI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   7800
         Picture         =   "frmOI.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdAbrir 
         BackColor       =   &H00C0C0C0&
         CausesValidation=   0   'False
         Height          =   450
         Left            =   6600
         Picture         =   "frmOI.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   450
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.CheckBox chkSelTodas 
         Alignment       =   1  'Right Justify
         Caption         =   "Seleccionar Todas"
         Height          =   255
         Left            =   6600
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskFechaIni 
         Height          =   345
         Left            =   1080
         TabIndex        =   7
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1095
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
         Left            =   2640
         TabIndex        =   8
         Tag             =   "Enc"
         ToolTipText     =   "Fecha del Movimiento"
         Top             =   1095
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
         Caption         =   "Producto:"
         Height          =   195
         Left            =   3960
         TabIndex        =   53
         Top             =   840
         Width           =   690
      End
      Begin VB.Label lblNombre 
         Height          =   375
         Left            =   2640
         TabIndex        =   27
         Top             =   1560
         Width           =   5655
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Concepto:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblAño1 
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1095
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "A:"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   1095
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame frTarjetas 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   8535
      Begin FPSpread.vaSpread spdDetalle 
         Height          =   1695
         Left            =   240
         OleObjectBlob   =   "frmOI.frx":0610
         TabIndex        =   56
         Top             =   240
         Width           =   8055
      End
      Begin VB.OptionButton mnOptTipo 
         Caption         =   "Todas"
         Height          =   375
         Index           =   4
         Left            =   4920
         TabIndex        =   46
         Top             =   2520
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton mnOptTipo 
         Caption         =   "No titulares"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   45
         Top             =   2520
         Width           =   1095
      End
      Begin VB.OptionButton mnOptTipo 
         Caption         =   "Titulares"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   2520
         Width           =   975
      End
      Begin VB.OptionButton mnOptTipo 
         Caption         =   "Adicionales"
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   43
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton mnOptTipo 
         Caption         =   "Reposiciones"
         Height          =   375
         Index           =   2
         Left            =   2400
         TabIndex        =   42
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox txtTotF 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   32
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   30
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7080
         TabIndex        =   28
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "IVA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   31
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "SubTotal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   29
         Top             =   2160
         Width           =   975
      End
   End
   Begin VB.Frame frmValores 
      Height          =   2655
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   8535
      Begin VB.Frame fraIva 
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   4695
         Begin VB.OptionButton OptIva 
            Caption         =   "0%"
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   50
            Top             =   160
            Width           =   975
         End
         Begin VB.OptionButton OptIva 
            Caption         =   "Exentos"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   49
            Top             =   160
            Width           =   1575
         End
         Begin VB.OptionButton OptIva 
            Caption         =   "Iva"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   160
            Width           =   1575
         End
      End
      Begin VB.TextBox txtNotasO 
         Height          =   1095
         Left            =   4800
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   720
         Width           =   3375
      End
      Begin VB.TextBox txtTotalO 
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
         Left            =   1560
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtIvaO 
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
         Left            =   1560
         TabIndex        =   20
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtSubtotalO 
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
         Left            =   1560
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Notas:"
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Total:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Iva:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Subtotal:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame frmPapel 
      Height          =   3375
      Left            =   240
      TabIndex        =   34
      Top             =   3240
      Width           =   8535
      Begin FPSpread.vaSpread spdFact 
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "frmOI.frx":099C
         TabIndex        =   57
         Top             =   240
         Width           =   8055
      End
      Begin VB.TextBox txtFact 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   36
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtTotalFact 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   35
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Facturas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   37
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.Label lblobs 
      AutoSize        =   -1  'True
      Caption         =   "Observaciones:"
      Height          =   195
      Left            =   240
      TabIndex        =   54
      Top             =   6720
      Width           =   1110
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Facturación Otros Ingresos"
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
      Left            =   4680
      TabIndex        =   12
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image Image3 
      Height          =   960
      Left            =   240
      Picture         =   "frmOI.frx":0DAD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   915
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      X1              =   1440
      X2              =   8760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "frmOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public impuesto As Double
Public TransAct As Boolean
Public TipoTarjeta As String
Public ConcepIva As Integer
Public prod As Byte


Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Facturacion;pwd=" & gpwdDataBase & ";database=" & gstrDataBase

If cboConcep.ItemData(cboConcep.ListIndex) = 11 Then
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptCheques.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Format(mskFechaIni, "yyyymmdd")
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaFin, "yyyymmdd")
    mdiMain.cryReport.StoredProcParam(2) = 0
ElseIf cboConcep.ItemData(cboConcep.ListIndex) = 12 Then
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptFacTarjetas.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = Val(txtCliente)
    mdiMain.cryReport.StoredProcParam(1) = Format(mskFechaIni, "yyyymmdd")
    mdiMain.cryReport.StoredProcParam(2) = Format(mskFechaFin, "yyyymmdd")
    prod = IIf(Product = 8, 6, Product)
    mdiMain.cryReport.StoredProcParam(3) = CStr(prod)

Else
    Exit Sub
End If


    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub


Private Sub cboBodegas_Click()
sqls = "select impuestointerior  iva from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly

If rsBD.EOF Then
    If Date >= "01/01/2010" Then
       impuesto = ivagral '0.15
    Else
       impuesto = 0.15
    End If
Else
    impuesto = rsBD!iva / 100
End If
rsBD.Close
Set rsBD = Nothing

If frmValores.Visible Then
    Call txtSubtotalO_Change
End If


If frTarjetas.Visible Then
    With spddetalle
     For i = 1 To .MaxRows
         .Row = i
         .Col = 4
         total = total + Val(.Text)
    
     Next i
    
    End With
    
    txtTotal.Text = total
    txtIva.Text = total * impuesto
    txtTotF.Text = total + txtIva
    
End If

End Sub

Private Sub cboConcep_click()
frTarjetas.Visible = False
frmPapel.Visible = False
frmValores.Visible = False
'fraIva.Visible = False
mskFechaIni = Date
mskFechaFin = Date
txtIvaO.Enabled = True
lblobs.Visible = False
txtobs.Visible = False
    
Select Case cboConcep.ItemData(cboConcep.ListIndex)
Case 11
    frmPapel.Visible = True
    mskFechaIni = Date - 1
    mskFechaFin = Date - 1
    Label3.Visible = False
    cboProductos.Visible = False
Case 12
    frTarjetas.Visible = True
    lblobs.Visible = True
    txtobs.Visible = True
'    mnOptTipo(3).Value = True
    Label3.Visible = True
    cboProductos.Visible = True
Case 13
    frmValores.Visible = True
    Label3.Visible = True
    cboProductos.Visible = True
Case 14
    fraIva.Visible = True
    frmValores.Visible = True
    OptIva(0).value = True
    txtIvaO.Enabled = False
    Label3.Visible = False
    cboProductos.Visible = False
Case Else
    frmValores.Visible = True
    Label3.Visible = False
    cboProductos.Visible = False
End Select

End Sub

Private Sub cboProductos_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & UCase(Trim(cboProductos.Text)) & " ','Leer'")
  If aqui <> Product Then
     InicializaForma
  End If
End Sub

Sub InicializaForma()
    lblNombre.Caption = ""
    txtCliente.Text = ""
End Sub



Private Sub cmdAbrir_Click()
If cboConcep.Text <> "" Then

Select Case cboConcep.ItemData(cboConcep.ListIndex)
    Case 11
        CargaSpreadPapel
    Case 12
        CargaSpreadTarjetas
End Select
Else
    MsgBox "Seleccione el concepto de la factura", vbInformation, "Facturacion OI"
    cboConcep.SetFocus
End If
End Sub

Private Sub cmdActualiza_Click()
Dim resp As String
Dim Folio As Long
Dim rsfol As Recordset
 
    Bodega = cboBodegas.ItemData(cboBodegas.ListIndex)
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 10 Then Bodega = 7
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 12 Then Bodega = 2
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 13 Then Bodega = 1
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 15 Then Bodega = 3
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 14 Then Bodega = 2
   If cboBodegas.ItemData(cboBodegas.ListIndex) = 11 Then Bodega = 3
   

   sqls = " select Prefijo serie, consecutivo " & _
         " From folios " & _
         " Where Bodega =" & Bodega & _
         " and tipo = 'FCM'"
   Set rsfol = New ADODB.Recordset
   rsfol.Open sqls, cnxBD, adOpenDynamic, adLockPessimistic
   
   If Not rsfol.EOF Then
      Folio = rsfol!consecutivo
   Else
      MsgBox "Error en folio de la Factura, verifique con sistemas", vbCritical
      Exit Sub
   End If
      

   resp = InputBox("Folio de Factura Nuevo", "Folio de Facturas de Comisiones", Folio)
   
  
If resp <> "" Then
   
  On Error GoTo err_gral

   sqls = "update folios set consecutivo = " & Val(resp) & _
          " , fechamodificacion = getdate()" & _
          " Where Bodega =" & Bodega & _
          " and tipo = 'FCM'"
   
   cnxBD.Execute sqls, intRegistros
   
   
End If
rsfol.Close
Set rsfol = Nothing
Exit Sub

err_gral:
         MsgBox ERR.Description, vbCritical, "Errores encontrados"
         Exit Sub
End Sub

Private Sub cmdBuscarC_Click()
Dim frmConsulta As New frmBusca_Cliente
If cboConcep.Text = "" Then
    MsgBox "Seleccione un concepto.", vbInformation
    Exit Sub
End If

If cboConcep.ItemData(cboConcep.ListIndex) = 12 Or cboConcep.ItemData(cboConcep.ListIndex) = 13 Then
    TipoBusqueda = "Cliente"
ElseIf cboConcep.ItemData(cboConcep.ListIndex) = 14 Or cboConcep.ItemData(cboConcep.ListIndex) = 16 Or cboConcep.ItemData(cboConcep.ListIndex) = 17 Or cboConcep.ItemData(cboConcep.ListIndex) = 18 Or cboConcep.ItemData(cboConcep.ListIndex) = 23 Then
    TipoBusqueda = "ClientesOI"
Else
    TipoBusqueda = "Emisores"
End If

frmConsulta.Show vbModal

If frmConsulta.cliente >= 0 Then
   txtCliente = frmConsulta.cliente
   lblNombre = frmConsulta.Nombre
End If
Set frmConsulta = Nothing
MsgBar "", False

End Sub

Private Sub cmdCancelar_Click()
Dim cliente, numero, tipo

If cboConcep.ItemData(cboConcep.ListIndex) = 11 Then
    With spdFact
          .Row = .ActiveRow
          .Col = 2
          Nombre = .Text
         
          If MsgBox("Desea quitar el pago a " & Nombre & "", vbQuestion + vbYesNo) = vbYes Then
              .Row = .ActiveRow
             .Action = 5
             .MaxRows = .MaxRows - 1
           
          End If
    End With

End If

End Sub

Private Sub cmdGrabar_Click()

    If cboConcep.ItemData(cboConcep.ListIndex) = 11 Then
        GrabaFacturaPapel
    ElseIf cboConcep.ItemData(cboConcep.ListIndex) = 13 Or cboConcep.ItemData(cboConcep.ListIndex) >= 15 Or cboConcep.ItemData(cboConcep.ListIndex) <= 23 Then
        GrabaFactura (cboConcep.ItemData(cboConcep.ListIndex))
    Else
        MsgBox "No se puede grabar factura. Favor de comunicarse a sistemas", vbCritical, "Error al grabar"
        Exit Sub
    End If
    
End Sub


Sub GrabaFacturaPapel()
Dim cliente As Integer, Folio As Long
Dim ArchImpre
Dim Comprobante As String
Dim Bodega As Integer


sTipoArch = "FACTURAS"
strPuerto = doFindPrinter(gstrPC, sTipoArch)
'SFileFact = "C:\facturacion\FactPapel.lis"
'nFileBat = FreeFile()
'sFileBat = "C:\Facturacion\FactOI.bat"

'Open sFileBat For Output As #nFileBat
'If gsOS = "XP" Then
'   Print #nFileBat, "PRINT /D:" & strPuerto & " " & SFileFact
'Else
'   Print #nFileBat, "COPY " & SFileFact & " " & strPuerto
'End If
'Close #nFileBat
  
'On Error GoTo err_gral
With spdFact
TransAct = False

     'operativo sorpresa (busca pifias en captura de entradas)
     For i = 1 To spdFact.MaxRows
        .Row = i
        .Col = 1
        cliente = CInt(.Text)
        .Col = 6
        folioR = CLng(.Text)

        sqls = "sp_Valida_entradas " & cliente & "," & folioR
        Set rsBD = New ADODB.Recordset
       rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
     
       If Not rsBD.EOF Then
         If Trim(rsBD!VALIDACION) = "INCORRECTO" Then
            MsgBox "El establecimiento " & cliente & " tiene un iva de comision que no corresponde a su sucursal" _
            & vbCrLf & "Favor de verificarlo!!!...", vbCritical, "Imposible seguir facturando"
            rsBD.Close
            Set rsBD = Nothing
            Exit Sub
         End If
       Else
         MsgBox "No hay establecimientos que validar", vbExclamation, "Sin establecimientos"
         rsBD.Close
         Set rsBD = Nothing
         Exit Sub
       End If
     Next
    '-----------------

For i = 1 To spdFact.MaxRows
    .Row = i
    .Col = 1
    cliente = CInt(.Text)
    .Col = 6
    folioR = CLng(.Text)
      
    sqls = " select a.sucursal, a.folio, a.cveestablecimiento cliente," & _
       " a.comision Importe, a.ivacomision Iva," & _
       " d.Descripcion DescCliente, d.rfc, d.Domicilio, d.Colonia, d.codigopostal, d.telefono, c.producto" & _
       " from entradaspagos a WITH (NOLOCK),establecimientos d , entradasdetpagos b, entradas c  " & _
       " Where d.Sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
       " and a.cveestablecimiento = " & cliente & " and a.folio = " & folioR & "  and a.fechapago between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'" & _
       " and a.status <= 1" & _
       " and a.comision > 0 " & _
       " and a.cveestablecimiento = d.cveestablecimiento" & _
       " and a.folio = b.folio" & _
       " and b.entrada = c.entrada" & _
       " and c.producto <> 5" & _
       " order by a.folio"


    Set rsBD = New ADODB.Recordset
    'rsbd.Open SQLS, cnxBD, adOpenDynamic, adLockReadOnly
    rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
    
    If Not rsBD.EOF Then
    
       
       cnxBD.BeginTrans
       TransAct = True
       Bodega = cboBodegas.ItemData(cboBodegas.ListIndex)
'       If Bodega = 10 Then Bodega = 7
'       If Bodega = 12 Then Bodega = 2
'       If Bodega = 11 Then Bodega = 3
'       If Bodega = 14 Then Bodega = 2
'       If Bodega = 15 Then Bodega = 9
'       If Bodega = 13 Then Bodega = 1
       sqls = " select Prefijo serie, consecutivo " & _
             " From folios " & _
             " Where Bodega =" & Bodega & _
             " and tipo = 'FCM'"
       Set rsfolio = New ADODB.Recordset
       'rsfolio.Open SQLS, cnxBD, adOpenDynamic, adLockPessimistic
       rsfolio.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
           
       If Not rsfolio.EOF Then
          sserie = Trim(rsfolio!Serie)
          Folio = rsfolio!consecutivo + 1
       Else
          cnxBD.RollbackTrans
          MsgBar "", False
          MsgBox "Falta valor de la serie ", vbOKOnly, "Avisar a Sistemas"
          Exit Sub
       End If
               
       If sserie = "" Then
          cnxBD.RollbackTrans
          MsgBar "", False
          MsgBox "Falta valor de la serie, Verifique con Sistemas  ", vbOKOnly, "Avisar a Sistemas"
          Exit Sub
       End If
       
    
    
        sqls = " EXEC sp_fm_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & rsBD!cliente
        sqls = sqls & vbCr & ", @Fecha        =  '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = " & cboConcep.ItemData(cboConcep.ListIndex)
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(sserie) & "'"
        sqls = sqls & vbCr & ", @Refer        = 2"
        sqls = sqls & vbCr & ", @Refer_Apl    = " & Folio
        sqls = sqls & vbCr & ", @CarAbo       = 'C'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & rsBD!importe
        sqls = sqls & vbCr & ", @iva      = " & rsBD!iva
        sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @TipoBon= " & rsBD!Producto
        '-SQLS = SQLS & vbCr & ", @Comprobante= " & rsBD!folio
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = '" & Usuario & "'"
        
        cnxBD.Execute sqls, intRegistros
    
        sqls = " exec sp_FM_facturas @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
               " ,@AnoFactura = " & Year(Date) & _
               " ,@Serie = '" & Trim(sserie) & "'" & _
               " ,@Factura = " & Folio & _
               " ,@Cliente = " & rsBD!cliente & _
               " ,@Fecha    = '" & Format(Date, "mm/dd/yyyy") & "' " & _
               " ,@Subtotal = " & rsBD!importe & _
               " ,@Iva    = " & rsBD!iva & _
               " ,@Rubro = " & cboConcep.ItemData(cboConcep.ListIndex) & _
               " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
               " ,@Status = 1" & _
               " ,@StatusImpreso = 0"
               
        cnxBD.Execute sqls, intRegistros
        
        
        sqls = " SELECT DISTINCT FACTURA FROM EntradasDet" & _
               " WHERE Entrada in(" & _
               " select entrada from entradasdetpagos " & _
               " where folio = " & rsBD!Folio & " )"
    
        Set rsdet = New ADODB.Recordset
        'rsdet.Open SQLS, cnxBD, adOpenDynamic, adLockReadOnly
        rsdet.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly
        Comprobante = " "
        
        Do While Not rsdet.EOF
            Comprobante = rsdet!Factura & " "
            rsdet.MoveNext
        Loop
        
        rsdet.Close
        Set rsdet = Nothing
        
               
        sqls = " exec sp_fm_facturas_detalle @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
               " ,@AnoFactura = " & Year(Date) & _
               " ,@Serie = '" & Trim(sserie) & "'" & _
               " ,@Factura = " & Folio & _
               " ,@Consecutivo   =1" & _
               " ,@Concepto = '" & Comprobante & "'" & _
               " ,@Cantidad = 1" & _
               " ,@PrecioVta = " & rsBD!importe & _
               " ,@PorcIva = " & rsBD!iva & _
               " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
        cnxBD.Execute sqls, intRegistros
        
        sqls = " update folios set consecutivo   = " & Folio & _
               " Where Bodega =" & Bodega & _
               " and tipo = 'FCM'"
             
        cnxBD.Execute sqls, intRegistros
        
        sqls = "update fm_clientes_movimientos set statusimpreso= 1" & _
               " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & "  and cliente = " & rsBD!cliente & _
               " and serie = '" & Trim(sserie) & "' and refer_Apl = " & Folio & _
               " and tipo_mov = " & cboConcep.ItemData(cboConcep.ListIndex) & " and refer =1"
               
        cnxBD.Execute sqls, intRegistros
        
        sqls = " update entradaspagos set status = 3 where  sucursal =" & rsBD!Sucursal & _
               " and folio = " & rsBD!Folio
        sqls = sqls & " and CveEstablecimiento=" & rsBD!cliente
               
        cnxBD.Execute sqls, intRegistros
           
        cnxBD.CommitTrans
        TransAct = False
        blnreimp = False
        
       ' Close #2
       ' ArchImpre = "C:\Facturacion\FactPapel.lis"
        'Open ArchImpre For Output As #2
        'Call Imprimefactura(cboBodegas.ItemData(cboBodegas.ListIndex), folio, Trim(sserie), cboConcep.ItemData(cboConcep.ListIndex))
        Call doGenArchFE_OI(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(sserie), Folio, Folio)
       
    End If
    Screen.MousePointer = 11
Next i
End With
Screen.MousePointer = 1
MsgBox "Facturas Generadas!!", vbInformation
cmdAbrir_Click
ImpUnaVez = False


   '---------------------
'
'      sTipoArch = "FACTURAS"
'      strPuerto = doFindPrinter(gstrPC, sTipoArch)
'      sFileFact = "C:\facturacion\FactPapel.lis"
'      nFileBat = FreeFile()
'      sFileBat = "C:\Facturacion\FactOI.bat"
'
'      Open sFileBat For Output As #nFileBat
'      If gsOS = "XP" Then
'         Print #nFileBat, "PRINT /D:" & strPuerto & " " & sFileFact
'      Else
'         Print #nFileBat, "COPY " & sFileFact & " " & strPuerto
'      End If
'      Close #nFileBat
'
'      EsperarShell (sFileBat)
'
'      DoEvents
'
'      Kill sFileFact
      
        
      '----------------------



'Close #14
'ArchImpre = "C:\Facturacion\FactOI.bat"
'Open ArchImpre For Output As #14
'
'SQLS = " SELECT impresora FROM BON_IMPRESORAS WHERE documento = 'FACTURAS' AND MAQUINA = " & gnMaquina
'Set rsBD = New ADODB.Recordset
'rsBD.Open SQLS, cnxBD, adOpenStatic, adLockReadOnly
'
''If rsBD.EOF Then
''  Print #14, "COPY  C:\facturacion\FactPapel.lis LPT1 "
''Else
''  Print #14, "copy c:\facturacion\FactPapel.lis " & Trim(rsBD!IMPRESORA)
''End If
'
'If rsBD.EOF Then
'    Print #14, "PRINT C:\facturacion\FactPapel.lis  "
'ElseIf Left(Trim(rsBD!Impresora), 3) = "LPT" Then
'    Print #14, "PRINT C:\facturacion\FactPapel.lis  "
'Else
'    'Print #14, "copy c:\facturacion\f" & nfol_fac & ".lis \\ARMANDO_P\IBM"
'    Print #14, "PRINT /D:" & Trim(rsBD!Impresora) & " c:\facturacion\FactPapel.lis "
'
'End If
'
'Close #14
'If MsgBox("Desea mandar la factura a la impresora", vbYesNo) = vbYes Then
'    Shell "C:\facturacion\factOI.BAT", 2
'End If
'
'Kill "C:\facturacion\FactPapel.lis  "
Exit Sub


rsBD.Close
Set rsBD = Nothing

Exit Sub


'err_gral:
'   MsgBox "Error " & Err.Number & ":" & Err.Description, , "Solicitud de Tarjetas"
'   If TransAct Then cnxBD.RollbackTrans
'   Call doErrorLog(gnBodega, "FACBE", Err.Number, Err.Description, Usuario, "frmoi.GrabaFacturaPapel")
'   MsgBar "", False

End Sub
Sub GrabaFactura(tipo As Integer)

Dim resp
Dim impuesto As Double
Dim rsimp As Recordset
Dim Folio As Long, Grupo As Integer
Dim ImporteT, ImporteF, Comision, Ivacom, TotComision As Double
Dim resultado
Dim Subtotal, iva, total

On Error GoTo err_gral

PRIMERA = True
TransAct = False

'cnxBD.BeginTrans

TransAct = True

Bodega = cboBodegas.ItemData(cboBodegas.ListIndex)

sqls = " select Prefijo serie, consecutivo " & _
      " From folios " & _
      " Where Bodega =" & Bodega & _
      " and tipo = 'FCM'"
Set rsBD = New ADODB.Recordset
'rsbd.Open SQLS, cnxBD, adOpenDynamic, adLockPessimistic
rsBD.Open sqls, cnxBD, adOpenDynamic, adLockOptimistic
    
If rsBD.EOF = False Then
   sserie = Trim(rsBD!Serie)
   Folio = rsBD!consecutivo + 1
Else
 '  cnxBD.RollbackTrans
   MsgBar "", False
   MsgBox "Falta valor de la serie ", vbOKOnly, "Avisar a Sistemas"
   Exit Sub
End If
rsBD.Close
Set rsBD = Nothing
        
sqls = "select impuestointerior  iva from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly

If rsBD.EOF Then
    If Date >= "01/01/2010" Then
       impuesto = ivagral * 100
    Else
       impuesto = 0.15
    End If
Else
    impuesto = rsBD!iva
End If
rsBD.Close
Set rsBD = Nothing


If sserie = "" Then
  ' cnxBD.RollbackTrans
   MsgBar "", False
   MsgBox "Falta valor de la serie, Verifique con Sistemas  ", vbOKOnly, "Avisar a Sistemas"
   Exit Sub
End If
 
If cboConcep.ItemData(cboConcep.ListIndex) = 12 Then
    Subtotal = txtTotal
    iva = txtIva
    'Total = txtTotal + txtIva
Else
    Subtotal = CDbl(txtSubtotalO)
    iva = CDbl(txtIvaO)
   ' Total = Subtotal + Iva
End If

If cboConcep.ItemData(cboConcep.ListIndex) <> 12 Then
    sqls = " EXEC sp_fm_Clientes_mov_ins "
    sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    sqls = sqls & vbCr & ", @Cliente      = " & Val(txtCliente)
    sqls = sqls & vbCr & ", @Fecha        = '" & Format(Date, "MM/DD/YYYY") & "'"
    sqls = sqls & vbCr & ", @Tipo_Mov     = " & tipo
    sqls = sqls & vbCr & ", @Serie        = '" & Trim(sserie) & "'"
    sqls = sqls & vbCr & ", @Refer        = 2"
    sqls = sqls & vbCr & ", @Refer_Apl    = " & Folio
    sqls = sqls & vbCr & ", @CarAbo       = 'C'"
    sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
    
    sqls = sqls & vbCr & ", @Importe      = " & Subtotal
    sqls = sqls & vbCr & ", @iva      = " & iva
    sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
    
    sqls = sqls & vbCr & ", @Fecha_Mov =  '" & Format(Date, "MM/DD/YYYY") & "'"
    sqls = sqls & vbCr & ", @Usuario = '" & gstrUsuario & "'"
    If cboConcep.ItemData(cboConcep.ListIndex) = 12 Or cboConcep.ItemData(cboConcep.ListIndex) = 13 Then
       sqls = sqls & vbCr & ", @TipoBon = " & Product
    End If
    
    If cboConcep.ItemData(cboConcep.ListIndex) = 14 Then
        sqls = sqls & vbCr & ", @Inversion = " & ConcepIva & ""
    End If
    
    cnxBD.Execute sqls, intRegistros
End If

sqls = " exec sp_FM_facturas @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
           " ,@AnoFactura = " & Year(Date) & _
           " ,@Serie = '" & Trim(sserie) & "'" & _
           " ,@Factura = " & Folio & _
           " ,@Cliente = " & Val(txtCliente) & _
           " ,@Fecha    = '" & Format(Date, "mm/dd/yyyy") & "' " & _
           " ,@Subtotal = " & Subtotal & _
           " ,@Iva    = " & iva & _
           " ,@Rubro = " & tipo & _
           " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
           " ,@Status = 1" & _
           " ,@StatusImpreso = 0"
           
cnxBD.Execute sqls, intRegistros

If tipo = 12 Or tipo = 13 Then
    sqls = " EXEC sp_Clientes_mov_ins "
    sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
    sqls = sqls & vbCr & ", @Cliente      = " & Val(txtCliente)
    sqls = sqls & vbCr & ", @Fecha        =    '" & Format(Date, "mm/dd/yyyy") & "'"
    sqls = sqls & vbCr & ", @Tipo_Mov     = " & tipo
    sqls = sqls & vbCr & ", @Serie        = '" & Trim(sserie) & "'"
    sqls = sqls & vbCr & ", @Refer        = 2"
    sqls = sqls & vbCr & ", @Refer_Apl    = " & Folio
    sqls = sqls & vbCr & ", @CarAbo       = 'C'"
    sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
    sqls = sqls & vbCr & ", @Importe      = " & CDbl(Subtotal) + CDbl(iva)
    sqls = sqls & vbCr & ", @Fecha_vento = '" & Format(Date + 1, "MM/DD/YYYY") & "'"
    sqls = sqls & vbCr & ", @Vendedor     = 0"
    sqls = sqls & vbCr & ", @CreditoFac = 'N'"
    sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
    sqls = sqls & vbCr & ", @Usuario = '" & Usuario & "'"
    sqls = sqls & vbCr & ", @TipoBon = " & Product
    
    cnxBD.Execute sqls, intRegistros

End If


If tipo = 12 Then

    sqls = " update solicitudesbe set FechaFac = getdate(), Factura = " & Folio & _
           " where tiposol = 1 and cliente =" & Val(txtCliente) & "  and status>=2  AND Producto=" & Product & _
           " and fecharesp between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'"
    sqls = sqls & " and Factura is null"
           
    If mnOptTipo(0).value = True Then
            sqls = sqls & " and tipo = 'T' "
    ElseIf mnOptTipo(1).value = True Then
            sqls = sqls & " and tipo = 'A' "
    ElseIf mnOptTipo(2).value = True Then
            sqls = sqls & " and tipo = 'RT' "
    ElseIf mnOptTipo(3).value = True Then
            sqls = sqls & " and tipo in  ('A', 'RT') "
    End If
                  
    cnxBD.Execute sqls, intRegistros


    With spddetalle
    
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        TipoTarjeta = .Text
        .Col = 2
        Cant = .Text
        .Col = 3
        costo = .Text
        
        If TipoTarjeta <> "" Then
        
        sqls = " exec sp_fm_facturas_detalle @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
               " ,@AnoFactura = " & Year(Date) & _
               " ,@Serie = '" & Trim(sserie) & "'" & _
               " ,@Factura = " & Folio & _
               " ,@Consecutivo   =" & i & _
               " ,@Concepto = '" & IIf(TipoTarjeta = "A", "ADICIONALES", IIf(TipoTarjeta = "R", "REPOSICIONES", IIf(TipoTarjeta = "T", "TITULARES", "X"))) & "'" & _
               " ,@Cantidad = " & Cant & _
               " ,@PrecioVta = " & costo & _
               " ,@PorcIva = " & Cant * costo * impuesto / 100 & _
               " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex)
        
        cnxBD.Execute sqls, intRegistros
        End If
    
    Next i
    End With
    

Else
  sqls = " exec sp_fm_facturas_detalle @Bodega =  " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
        " ,@AnoFactura = " & Year(Date) & _
        " ,@Serie = '" & Trim(sserie) & "'" & _
        " ,@Factura = " & Folio & _
        " ,@Consecutivo   =1 " & _
        " ,@Concepto = '" & Left(UCase(Trim(txtNotasO)), 400) & "'" & _
        " ,@Cantidad = 1" & _
        " ,@PrecioVta = " & Subtotal & _
        " ,@PorcIva = " & iva & _
        " ,@BodegaOrigen =  " & cboBodegas.ItemData(cboBodegas.ListIndex)
    
    cnxBD.Execute sqls, intRegistros

End If


sqls = " update folios set consecutivo   = " & Folio & _
  " Where Bodega =" & Bodega & _
  " and tipo = 'FCM'"

cnxBD.Execute sqls, intRegistros

    
'cnxBD.CommitTrans
TransAct = False

'Call Imprimefactura(cboBodegas.ItemData(cboBodegas.ListIndex), folio, Trim(sserie), Tipo)

Observ = Trim(UCase(txtobs.Text))
Call doGenArchFE_OI(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(sserie), Folio, Folio)
'----------------
ImpUnaVez = False
MsgBox "Factura " & Trim(sserie) & Folio & " generada!"
    
Exit Sub

err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmOI.GrabaFactura")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Solicitud de Tarjetas"
   If TransAct Then Resume Next 'cnxBD.RollbackTrans
   Resume Next
   MsgBar "", False

End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Imprime (crptToWindow)
End Sub

Private Sub Command2_Click()
'Call doGenArchFE_OI(7, "OG", 184, 184)
frmFact.Show 1
 
 'Call doGenArchFE_OI(1, "OA", 15695, 15695)
End Sub

Private Sub Form_Load()
    CargaBodegas cboBodegas
    CargaConcep cboConcep
    Observ = ""
    mskFechaIni = Date
    mskFechaFin = Date
    frTarjetas.Visible = False
    frmValores.Visible = False
    lblobs.Visible = False
    txtobs.Visible = False
    ImpUnaVez = False
    Command1.Enabled = False
    cmdCancelar.Enabled = False
    Label3.Visible = False
    cboProductos.Visible = False
    Call CargaComboBE(cboProductos, "sp_sel_productobe 'BE','','Cargar'")
    Call LeeproductoBE(cboProductos, "sp_sel_productobe 'BE','" & Trim(cboProductos.Text) & " ','Leer'")
'    cboProductos.Text = UCase("Winko Mart")
        
    'SQLS = "exec  sp_AsigPapFact_Sel " & Val(sUbicacion) & ", '" & gstrImpFact & "', 'OI', 1,1"
    'Set rsbd = New ADODB.Recordset
    'rsbd.Open SQLS, cnxBD, adOpenDynamic, adLockReadOnly
    If Dir("C:\Facturacion\Paso", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema de facturaciòn en este equipo, favor de comunicarse a sistemas"
        cmdAbrir.Enabled = False
        cmdGrabar.Enabled = False
    End If
    If Dir("C:\Facturacion\Pruebas", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema de facturaciòn en este equipo, favor de comunicarse a sistemas"
        cmdAbrir.Enabled = False
        cmdGrabar.Enabled = False
    End If
    If Dir("C:\GoDir", vbDirectory) = "" Then
        MsgBox "No tiene configurado el sistema GOSOFT, favor de comunicarse a sistemas"
        cmdAbrir.Enabled = False
        cmdGrabar.Enabled = False
    End If
End Sub

Sub CargaSpreadPapel()
Dim Nombre As String
Dim cliente As Long
Dim total  As Double
Dim TotalF As Long

On Error GoTo err_gral

sqls = " select distinct a.sucursal, a.folio, a.cveestablecimiento cliente," & _
       " a.comision Importe, a.ivacomision Iva," & _
       " d.Descripcion DescCliente, d.rfc, d.Domicilio, d.Colonia, d.codigopostal, d.telefono, c.producto" & _
       " from entradaspagos a WITH (NOLOCK),establecimientos d , entradasdetpagos b, entradas c   " & _
       " Where d.Sucursal = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
       " and a.fechapago between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & "'" & _
       " and a.status < = 1" & _
       " and a.comision > 0 " & _
       " and a.cveestablecimiento = d.cveestablecimiento" & _
       " and a.folio = b.folio" & _
       " and b.entrada = c.entrada" & _
       " and c.producto <> 5" & _
       " AND a.Tipopago=b.Tipopago" & _
       " order by a.folio"



'SQLS = "exec spr_Impresionfacturas @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
        ",@FechaIni = '" & Format(mskFechaIni, "mm/dd/yyyy") & "'" & _
        ",@FechaFin = '" & Format(mskFechaFin, "mm/dd/yyyy") & "'"

Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly


If rsBD.EOF Then
    MsgBox "No hay facturas por generar", vbInformation, "Facturacion OI"
    spdFact.Col = -1
    spdFact.Row = -1
    spdFact.Action = 12
    spdFact.MaxRows = 0
    txtFact = 0
    txtTotalFact = 0
    Exit Sub
End If
   
With spdFact
   .Col = -1
   .Row = -1
   .Action = 12
   .MaxRows = 0
   i = 1
   total = 0

   Do While Not rsBD.EOF
      .MaxRows = i
      .Row = i
      .Col = 1
'      .Text = rsBD!Factura
 '     .Col = 2
      cliente = rsBD!cliente
      .Text = cliente
      .Col = 2
      .Text = rsBD!DescCliente
      .Col = 3
      .Text = CDbl(rsBD!importe)
      .Col = 4
      .Text = CDbl(rsBD!iva)
      .Col = 5
      .Text = CDbl(rsBD!importe) + CDbl(rsBD!iva)
      .Col = 6
      .Text = CLng(rsBD!Folio)
      .Col = 7
      .Text = Val(rsBD!Producto)
      TotalF = TotalF + 1
      total = total + CDbl(rsBD!importe) + CDbl(rsBD!iva)
      
      i = i + 1
      rsBD.MoveNext
   Loop

   txtFact = TotalF
   txtTotalFact = total
   

End With

rsBD.Close
Set rsBD = Nothing

Exit Sub
 
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmOI.CargaSpreadPapel")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Solicitud de Tarjetas"
   MsgBar "", False

End Sub

Sub CargaSpreadTarjetas()
Dim Nombre As String
Dim cliente As Long
Dim total  As Double

On Error GoTo err_gral

If Val(txtCliente) = 0 Then
    MsgBox "Primero capture el cliente!", vbInformation, "Facturacion OI"
    Exit Sub
End If
    prod = Product
    sqls = "select a.cliente , left(a.tipo,1) Tipo ,  count(a.cliente)  Cant, b.costo,  count(a.cliente) *  b.costo Total" & _
       " from solicitudesbe a with (nolock), costostar b" & _
       " where a.cliente =" & Val(txtCliente) & " and a.Factura is null" & _
       " and a.fechasol between '" & Format(mskFechaIni, "mm/dd/yyyy") & "' and '" & Format(mskFechaFin, "mm/dd/yyyy") & " 23:59:00'" & _
    " and a.Producto=" & prod
     If mnOptTipo(0).value = True Then
             sqls = sqls & " and a.tipo = 'T' "
     ElseIf mnOptTipo(1).value = True Then
             sqls = sqls & " and a.tipo = 'A' "
     ElseIf mnOptTipo(2).value = True Then
             sqls = sqls & " and a.tipo in ('RT', 'RA') "
     ElseIf mnOptTipo(3).value = True Then
             sqls = sqls & " and a.tipo in  ('A', 'RT', 'RA') "
     ElseIf mnOptTipo(3).value = True Then
             sqls = sqls & " and a.tipo in  ('A', 'RT', 'RA','T') "
     End If
       sqls = sqls & " and a.cliente = b.cliente" & _
       " and left(a.tipo,1)= b.tipo  " & _
       " and b.costo >0 " & _
       " group by a.cliente, left(a.tipo,1), b.costo"

 Set rsBD2 = New ADODB.Recordset
 rsBD2.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly


If rsBD2.EOF Then
    MsgBox "No hay tarjetas por facturar para este cliente", vbInformation, "Facturacion OI"
    
   spddetalle.Col = -1
   spddetalle.Row = -1
   spddetalle.Action = 12
   spddetalle.MaxRows = 0
 
   ' txtCliente = ""
    'lblNombre = ""
    
    Exit Sub
End If
   
 With spddetalle
   .Col = -1
   .Row = -1
   .Action = 12
   .MaxRows = 0
   i = 1
   total = 0

   Do While Not rsBD2.EOF
      .MaxRows = i
      .Row = i
      .Col = 1
      cliente = rsBD2!cliente
      .Text = cliente
      .Col = 2
      Nombre = BuscaCliente(cliente, Nombre)
      .Text = Nombre
      .Col = 1
      .Text = Trim(rsBD2!tipo)
      .Col = 2
      .Text = CInt(rsBD2!Cant)
      .Col = 3
      .Text = CDbl(rsBD2!costo)
      .Col = 4
      .Text = CDbl(rsBD2!total)
      total = total + CDbl(rsBD2!total)
      i = i + 1
      rsBD2.MoveNext
   Loop

   txtTotal = total
   txtIva = total * impuesto
   txtTotF = Format(Val(txtTotal) + Val(txtIva), "###,###,###.00")

End With

rsBD2.Close
Set rsBD2 = Nothing

Exit Sub
 
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmOI.CargaSpreadTarjetas")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Solicitud de Tarjetas"
   MsgBar "", False
End Sub


Private Sub Label15_Click()

End Sub



Private Sub mnOptTipo_Click(Index As Integer)
    Select Case Index
        Case 0
            TipoTarjeta = "T"
        Case 1
            TipoTarjeta = "A"
        Case 2
            TipoTarjeta = "R"
        Case 3
            TipoTarjeta = "TT"
    End Select
            
End Sub

Private Sub OptIva_Click(Index As Integer)

ConcepIva = Index
Select Case Index
    Case 0 'iva
        txtIvaO = Format(Val(txtSubtotalO) * impuesto, "#########.00")
        
    Case 1, 2
        txtIvaO = 0
End Select

 txtTotalO = CDbl(Val(txtSubtotalO)) + CDbl(Val(txtIvaO))
    
End Sub

Private Sub spdDetalle_Change(ByVal Col As Long, ByVal Row As Long)
Dim total As Double

With spddetalle

If Col = 3 Or Col = 2 Then

    For i = 1 To .MaxRows
        .Row = i
        .Col = 4
        total = total + Val(.Text)

    Next i

End If
End With

txtTotal.Text = total

End Sub

Private Sub spddetalle_KeyPress(KeyAscii As Integer)
With spddetalle
 If KeyAscii = 13 Then
    If .ActiveCol = 2 Or .ActiveCol = 3 Then
        .Row = .ActiveRow
        .Col = 2
        Cant = Val(.Text)
        .Col = 3
        valor = Val(.Text)
        total = Cant * valor
        .Col = 4
        .Text = total
    End If
 End If
End With

End Sub

Private Sub txtcliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"

End Sub

Private Sub txtCliente_LostFocus()

If txtCliente = "" Then Exit Sub

Select Case cboConcep.ItemData(cboConcep.ListIndex)

Case 14, 16, 17, 18, 22

    sqls = " select nombre  from clientesoi " & _
            " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
            " and cliente = " & Val(txtCliente)


Case Else

    sqls = " select nombre  from clientes " & _
            " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex) & _
            " and cliente = " & Val(txtCliente)
            
End Select
            
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenDynamic, adLockReadOnly

If rsBD.EOF Then
    MsgBox "Cliente no existe o no corresponde a la sucursal de " & cboBodegas.Text
    txtCliente = ""
    txtCliente.SetFocus
       
    Exit Sub
Else
    lblNombre = rsBD!Nombre
End If

End Sub

Private Sub txtIvaO_Change()
  txtTotalO = CDbl(Val(txtSubtotalO)) + CDbl(Val(txtIvaO))
End Sub

Private Sub txtSubtotalO_Change()
    If cboConcep.ItemData(cboConcep.ListIndex) = 15 Then
        txtIvaO = 0
    Else
        txtIvaO = Format(Val(txtSubtotalO) * impuesto, "#########.00")
    End If
    txtTotalO = CDbl(Val(txtSubtotalO)) + CDbl(Val(txtIvaO))
    
    If cboConcep.ItemData(cboConcep.ListIndex) = 14 Then
        If OptIva(1).value = True Or OptIva(2).value = True Then
             txtIvaO = 0
              txtTotalO = CDbl(Val(txtSubtotalO)) + CDbl(Val(txtIvaO))
        End If
    End If

End Sub

Private Sub txtTotal_Change()
'    SQLS = "select impuestointerior from bodegas where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
'
'    Set rsbd = New ADODB.Recordset
'    rsbd.Open SQLS, cnxBD, adOpenDynamic, adLockReadOnly
'
    
    
  '  If Not rsbd.EOF Then
        txtIva.Text = CDbl(txtTotal.Text) * impuesto
   ' Else
    '    txtIva.Text = CDbl(txtTotal.Text) * 0.15
    'End If
        
    txtTotF = CDbl(txtTotal.Text) + CDbl(txtIva.Text)
    
End Sub
