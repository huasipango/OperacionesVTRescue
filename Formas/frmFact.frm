VERSION 5.00
Begin VB.Form frmFact 
   Caption         =   "Factura"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   3300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   495
      Left            =   1080
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox txtfactura 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtserie 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtbodega 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtproducto 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Serie:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bodega:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   690
   End
End
Attribute VB_Name = "frmFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Dim Bodega As Byte, Serie As String, Factura As Double
     Bodega = Val(txtbodega.Text)
     Serie = Trim(UCase(txtSerie.Text))
     Factura = Val(txtFactura.Text)
     Producto_factura = Val(txtproducto.Text) '6
     Product = Val(txtproducto.Text)
     Call Producto_actual
     If Reimp_FCNC = 3 Then
        Call doGenArchFE_NCA(Bodega, Serie, Factura, Factura, Reimp_FCNC)
     ElseIf Mid(Serie, 1, 1) = "V" Then
        Call doGenArchFE_OI(Bodega, Serie, Factura, Factura) 'ESTE ES EL DE OI
     Else
        Call doGenArchFE(Bodega, Serie, Factura, Factura, Reimp_FCNC)
       ' Call doGenArchFE(5, "CE", 70, 70, 3)
     End If
End Sub
