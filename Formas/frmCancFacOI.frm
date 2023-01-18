VERSION 5.00
Begin VB.Form frmCancFacOI 
   Caption         =   "Cancelación de Facturas OI"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   4920
      TabIndex        =   13
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cancelar"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   570
         Left            =   240
         Picture         =   "frmCancFacOI.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Salir"
         CausesValidation=   0   'False
         Height          =   570
         Left            =   240
         Picture         =   "frmCancFacOI.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
   End
   Begin VB.Frame frCte 
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtTipoMov 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox txtValor 
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
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtFechaFac 
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
         TabIndex        =   4
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtSerie 
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
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Width           =   375
      End
      Begin VB.ComboBox cboBodegas 
         Height          =   315
         ItemData        =   "frmCancFacOI.frx":0204
         Left            =   1440
         List            =   "frmCancFacOI.frx":020B
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2775
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
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1095
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
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   855
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
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
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
         Left            =   3120
         TabIndex        =   11
         Top             =   720
         Width           =   615
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
         TabIndex        =   10
         Top             =   240
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   720
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
         Height          =   735
         Left            =   2760
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCancFacOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
        
        
On Error GoTo err_gral

        
        Set rsBD = Nothing
        
        'cnxBD.BeginTrans

        sqls = " EXEC sp_fm_Clientes_mov_ins "
        sqls = sqls & vbCr & "  @Bodega       = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & vbCr & ", @Cliente      = " & txtCliente
        sqls = sqls & vbCr & ", @Fecha        =    '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Tipo_Mov     = 70"
        sqls = sqls & vbCr & ", @Serie        = '" & Trim(txtSerie) & "'"  'Serie Nota Credito
        sqls = sqls & vbCr & ", @Refer        =1"
        sqls = sqls & vbCr & ", @Refer_Apl    = " & txtFactura
        sqls = sqls & vbCr & ", @CarAbo       = 'A'"
        sqls = sqls & vbCr & ", @Tipo_Mov_Apl = 10"
        sqls = sqls & vbCr & ", @Importe      = " & txtValor
        sqls = sqls & vbCr & ", @Fecha_Mov = '" & Format(Date, "MM/DD/YYYY") & "'"
        sqls = sqls & vbCr & ", @Usuario = " & Usuario
        
        
        cnxBD.Execute sqls, intRegistros
        
        If txtTipoMov.Text = 12 Or txtTipoMov.Text = 13 Then
        
            sqls = " EXEC sp_Clientes_mov_ins @Bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
            sqls = sqls & vbCr & " , @Cliente      = " & txtCliente
            sqls = sqls & vbCr & " , @Fecha        =  '" & Format(Date, "MM/DD/YYYY") & "'"
            sqls = sqls & vbCr & " , @Tipo_Mov     = 88"
            sqls = sqls & vbCr & " , @Serie        = '" & Trim(txtSerie) & "'"
            sqls = sqls & vbCr & " , @Refer        = 1"
            sqls = sqls & vbCr & " , @Refer_Apl    = " & txtFactura
            sqls = sqls & vbCr & " , @CarAbo       = 'A'"
            sqls = sqls & vbCr & " , @Tipo_Mov_Apl = 10"
            sqls = sqls & vbCr & " , @Importe      = " & txtValor
            sqls = sqls & vbCr & ",  @Fecha_Mov ='" & Format(Date, "MM/DD/YYYY") & "'"
            sqls = sqls & vbCr & ",  @Usuario = " & Usuario
        
            cnxBD.Execute sqls, intRegistros
        
        End If
        
        sqls = " update FM_FACTURAS "
        sqls = sqls & " set status = 2"
        sqls = sqls & " where bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & " and factura= " & Val(txtFactura)
        cnxBD.Execute sqls, intRegistros
        
        Call doGenArchCanc(cboBodegas.ItemData(cboBodegas.ListIndex), Trim(txtSerie.Text), txtFactura.Text, 4)
        
        'cnxBD.RollbackTrans
        
        MsgBox "Factura " & txtFactura & " Cancelada!", vbInformation, "Cancelación de Facturas OI"
        LimpiarControles Me
        CargaBodegas cboBodegas
        
        Exit Sub
        
err_gral:
   Call doErrorLog(gnBodega, "OPE", ERR.Number, ERR.Description, Usuario, "frmCancFacOI.cmdCancelar_Click")
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Cancelar Nota"
   MsgBar "", False
    
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    LimpiarControles Me
    CargaBodegas cboBodegas
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
If ValidaNumericos(KeyAscii, txtFactura, 3) Then entertab (KeyAscii)
    If KeyAscii = 13 Then
      If Trim(txtFactura) <> "" Then
        sqls = "select *  "
        sqls = sqls & " from fm_clientes_movimientos a"
        sqls = sqls & " Where a.bodega = " & cboBodegas.ItemData(cboBodegas.ListIndex)
        sqls = sqls & " and a.refer_apl= " & Val(txtFactura)
        sqls = sqls & " AND Fecha>='01/01/2011' order by tipo_mov"
        
        
        Set rsBD = New ADODB.Recordset
        rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

        If Not rsBD.EOF Then
            txtCliente = rsBD!Cliente
            'txtFactura = rsBD!refer_apl
            txtSerie = rsBD!serie
            txtFechaFac = rsBD!Fecha
            txtValor = Format(rsBD!Importe + rsBD!iva, "#######0.00")
            txtTipoMov = rsBD!TIPO_MOV
            If Format(CDate(rsBD!Fecha), "MM/DD/YYYY") <> Format(Date, "MM/DD/YYYY") Then
                MsgBox "No se puede cancelar la factura por la fecha, solo se pueden cancelar las facturas del dia de hoy.", vbCritical, "Cancelación de Facturas"
                cmdCancelar.Enabled = False
            Else
                cmdCancelar.Enabled = True
            End If
        Else
            MsgBox "Factura no existe, favor de verificarla!", vbCritical, "Cancelación de Facturas"
            txtCliente = ""
            txtFactura = 0
            txtSerie = ""
            txtFechaFac = ""
            txtValor = 0
            
        End If
        rsBD.Close
        Set rsBD = Nothing
    End If
   End If
          
    
End Sub

