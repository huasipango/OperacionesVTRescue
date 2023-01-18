VERSION 5.00
Begin VB.Form frmCancNotasCargo 
   Caption         =   "Cancelación de Notas de Cargo"
   ClientHeight    =   3975
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   7335
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Salir"
         CausesValidation=   0   'False
         Height          =   570
         Left            =   3960
         Picture         =   "frmCancNotasCargo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Height          =   570
         Left            =   2880
         Picture         =   "frmCancNotasCargo.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame frCte 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmCancNotasCargo.frx":0204
         Left            =   1440
         List            =   "frmCancNotasCargo.frx":020B
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtFolio 
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
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFactura 
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
         Left            =   1440
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1440
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   3480
         TabIndex        =   7
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtFechaFac 
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
         Left            =   5640
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtFechaNota 
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
         Left            =   1440
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtValor 
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
         Left            =   3480
         TabIndex        =   4
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtSerieNota 
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
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   375
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
         TabIndex        =   21
         Top             =   360
         Width           =   975
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
         Left            =   2640
         TabIndex        =   16
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "Folio:"
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
         TabIndex        =   15
         Top             =   840
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
         TabIndex        =   14
         Top             =   1800
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
         Top             =   1320
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
         Left            =   2760
         TabIndex        =   12
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Factura:"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Nota:"
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
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Valor:"
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
         Left            =   2760
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCancNotasCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
       
        
On Error GoTo err_gral

    If Format(CDate(txtFechaNota.Text), "MM/DD/YYYY") = Format(Date, "MM/DD/YYYY") Then
        
        Set rsBD = Nothing
        
     
        sqls = " EXEC sp_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & txtCliente
        sqls = sqls & vbCr & ", @Fecha        =    '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = 83"
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(txtserie) & "'"
        sqls = sqls & vbCr & ", @Refer        = " & txtFolio
        sqls = sqls & vbCr & ", @Refer_Apl    = " & txtFactura
        sqls = sqls & vbCr & ", @CarAbo       = 'C'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Cuenta_origen = " & txtFolio
        sqls = sqls & vbCr & ", @Importe      = " & txtValor
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = " & Usuario
        
        
        cnxBD.Execute sqls, intRegistros
        
        
        sqls = " update notascar "
        sqls = sqls & " set status = 2"
        sqls = sqls & " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & " and folio= " & Val(txtFolio)
        cnxBD.Execute sqls, intRegistros
        
        
        Call doGenArchCanc(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(txtSerieNota.Text), Val(txtFolio), 3)
       
        MsgBox "¡Nota " & txtFolio & " Cancelada!", vbInformation, "Cancelación de Notas de Cargo"
        LimpiarControles Me
        CargaBodegas cboBodegas
    Else
        MsgBox "No se pueden cancelar notas, despues del día que se realizaron", vbCritical
    
    End If
        
        Exit Sub
        
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmCancNotas.cmdCancelar_Click")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Cancelar Nota"
   Resume Next
   MsgBar "", False
    
    
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    txtSerieNota.Text = "NC"
    LimpiarControles Me
    CargaBodegas cboBodegas
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)

If ValidaNumericos(KeyAscii, txtFolio, 3) Then entertab (KeyAscii)
    If KeyAscii = 13 Then
      If Trim(txtFolio) <> "" Then
        sqls = "select a.bodega, a.cliente, a.serie,b.serienota SerieNota, a.refer_apl,a.fecha, b.fechaemi, b.valor"
        sqls = sqls & " from clientes_movimientos a, notascar b"
        sqls = sqls & " Where a.bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & " and a.Tipo_mov = 43 "
        sqls = sqls & " and a.refer= " & Val(txtFolio)
        sqls = sqls & " and  a.refer = b.folio and b.serienota = 'NC'"
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

        If Not rsBD.EOF Then
            txtCliente = rsBD!cliente
            txtFactura = rsBD!refer_apl
            txtserie = rsBD!serie
            txtSerieNota = rsBD!serieNota
            txtFechaFac = rsBD!Fecha
            txtFechaNota = rsBD!fechaemi
            txtValor = Format(rsBD!valor, "#######0.00")
            
            
        Else
            MsgBox "Folio no existe, favor de verificarlo!", vbCritical, "Cancelación de Notas"
        End If
        rsBD.Close
        Set rsBD = Nothing
    End If
   End If
          
End Sub
