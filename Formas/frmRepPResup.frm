VERSION 5.00
Begin VB.Form frmRepPResup 
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   2280
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Producto"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   4095
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
         ItemData        =   "frmRepPResup.frx":0000
         Left            =   120
         List            =   "frmRepPResup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3480
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtIngXtarjeta 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtTasaProm 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtInvProm 
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Ingresos por Tarjeta:"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Tasa Prom de Inv Mes:"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Inversion Promedio en el mes:"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdPresentar 
      Height          =   450
      Left            =   4440
      Picture         =   "frmRepPResup.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Det"
      ToolTipText     =   "Imprime en la Pantalla"
      Top             =   1080
      Width           =   450
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Height          =   450
      Left            =   4440
      Picture         =   "frmRepPResup.frx":011E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   1680
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   4095
      Begin VB.ComboBox cboAño 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRepPResup.frx":0290
         Left            =   2520
         List            =   "frmRepPResup.frx":0292
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmRepPResup.frx":0294
         Left            =   1200
         List            =   "frmRepPResup.frx":02BF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Mes:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmRepPResup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mclsAniform As clsAnimated
Attribute mclsAniform.VB_VarHelpID = -1
Dim clsA As New clsAnimated
Dim prod As Byte
Dim aant As Double

Private Sub cboProducto_Click()
Dim aqui As Byte
  aqui = Product
  Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & UCase(Trim(cboProducto.Text)) & " ','Leer'")
  If aqui <> Product Then
     'InicializaForma
  End If
End Sub

Sub CalculaVentas()
Dim rsvta As ADODB.Recordset, rsvtaAA As ADODB.Recordset, rsPres As ADODB.Recordset
Dim pase As Double

fechaini = Format(cboMes.ListIndex + 1, "00") + "/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")

FechaIniAA = Format(cboMes.ListIndex + 1, "00") + "/01/" + CStr(Val(cboAño.Text) - 1)
FechaFinAA = FechaFinMes(cboMes.ItemData(cboMes.ListIndex), CStr(Val(cboAño.Text) - 1))

If Product = 1 Then
 sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
  sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(BON_FAC_BONGRA)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa = " & cboMes.ListIndex + 1 & _
       " and producto = " & Product

Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

 
sqls = "select * from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    ventapresupacum = 0
    MsgBox "No esta capturado el presupuesto para este periodo"
    PorcSVta = IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)
    If PorcSVta < 0 Then
        PorcSVta = 0
    Else
        PorcSVta = 100
    End If
    Presup = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) / 1000
Else
    ventapresupacum = rsPres!Ventas
    PorcSVta = IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(rsPres!Ventas), 0, rsPres!Ventas)
    If PorcSVta < 0 Then
        PorcSVta = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) / IIf(IsNull(rsPres!Ventas), 0, rsPres!Ventas)) * 100
    Else
        PorcSVta = 100
    End If
    Presup = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(rsPres!Ventas), 0, rsPres!Ventas)) / 1000
End If




Venta = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) / 1000
aant = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(rsvtaAA!Ventas), 0, rsvtaAA!Ventas)) / 1000
pase = IIf(IsNull(rsvtaAA!Ventas), 0, rsvtaAA!Ventas) * 100
pase = IIf(pase = 0, 1, pase)
seaa = (((IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) / pase) - PorcSVta)
 
fechaini = "01/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")


If Product = 1 Then
   sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
   sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(BON_FAC_BONGRA)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa <= " & cboMes.ListIndex + 1 & _
       " and producto = " & Product
Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

 
sqls = "select sum(ventas) Ventas from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo between  " & Val(cboAño.Text) & "01" & " and " & Val(cboAño.Text) & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    ventapresupacum = 0
    MsgBox "No esta capturado el presupuesto para el periodo del año anterior"
    If PorcSVtaAcum < 0 Then
        PorcSVtaAcum = 0
    Else
        PorcSVtaAcum = 100
    End If
Else
    ventapresupacum = rsPres!Ventas
    If PorcSVtaAcum < 0 Then
        PorcSVtaAcum = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) / IIf(IsNull(rsPres!Ventas), 0, rsPres!Ventas)) * 100
    Else
        PorcSVtaAcum = 100
    End If
End If

PorcSVtaAcum = IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(ventapresupacum), 0, ventapresupacum)

VentaAcum = IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) / 1000
PresupAcum = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(ventapresupacum), 0, ventapresupacum)) / 1000
AAntAcum = (IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(rsvtaAA!Ventas), 0, rsvtaAA!Ventas)) / 1000
seaaacum = (((IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) / IIf(IsNull(rsvtaAA!Ventas), 1, rsvtaAA!Ventas)) * 100) - PorcSVta


sqls = " exec sp_EdoResultados  @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
       " , @Concepto = 1" & _
       " , @Mreal = " & Format(Venta, "#########.00") & _
       " , @Porcsvta = " & Format(PorcSVta, "#########.00") & _
       " , @Presup =  " & Format(Presup, "#########.00") & _
       " , @aant = " & Format(aant, "#########.00") & _
       " , @seaa = " & Format(se < aa, "#########.00") & _
       " , @acumreal = " & Format(VentaAcum, "#########.00") & _
       " , @acumporcsvta = " & Format(PorcSVtaAcum, "#########.00") & _
       " , @acumpresup = " & Format(PresupAcum, "#########.00") & _
       " , @acumaant= " & Format(AAntAcum, "#########.00") & _
       " , @acumseaa = " & Format(seaaacum, "########.00")

cnxBD.Execute sqls, intRegistros

End Sub

Sub CalculaIngxTarjetas()
Dim rsvta As ADODB.Recordset, rsvtaAA As ADODB.Recordset, rsPres As ADODB.Recordset

fechaini = Format(cboMes.ListIndex + 1, "00") + "/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")

FechaIniAA = Format(cboMes.ListIndex + 1, "00") + "/01/" + CStr(Val(cboAño.Text) - 1)
FechaFinAA = FechaFinMes(cboMes.ItemData(cboMes.ListIndex), CStr(Val(cboAño.Text) - 1))

If Product = 1 Then
   sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
    sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(BON_FAC_BONGRA)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa = " & cboMes.ListIndex + 1 & _
       " and producto = " & Product

Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = "select ingxtarjeta from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    IngPresup = 0
    MsgBox "No esta capturado el presupuesto para este periodo", vbInformation, "Presupuesto no capturado"
Else
    IngPresup = rsPres!ingxtarjeta
End If

sqls = "select Importe=IsNull(Sum(Case" & _
              " When (CarAbo='C' )" & _
              " Then Importe " & _
              " When (CarAbo='A' ) " & _
              " Then Importe*-1 " & _
              " End),0) " & _
"From clientes_movimientos" & _
 " where fecha between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
 " and tipo_mov in (12,13,88)"
 

Set rsing = New ADODB.Recordset
rsing.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = "select sum(importe) Importe from clientes_movimientos" & _
       " where fecha between '" & FechaIniAA & "' and '" & FechaFinAA & " 23:59:00'" & _
      " and tipo_mov in (12,13)"
 
 
sqls = "select Importe=IsNull(Sum(Case" & _
              " When (CarAbo='C' )" & _
              " Then Importe " & _
              " When (CarAbo='A' ) " & _
              " Then Importe*-1 " & _
              " End),0) " & _
"From clientes_movimientos" & _
" where fecha between '" & FechaIniAA & "' and '" & FechaFinAA & " 23:59:00'" & _
" and tipo_mov in (12,13,88)"
 

Set rsingaa = New ADODB.Recordset
rsingaa.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


Venta = (IIf(IsNull(rsing!importe), 0, rsing!importe)) / 1000
PorcSVta = (IIf(IsNull(rsing!importe), 0, rsing!importe) / IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) * 100
Presup = (IIf(IsNull(rsing!importe), 0, rsing!importe) - IngPresup) / 1000
aant = (IIf(IsNull(rsing!importe), 0, rsing!importe) - IIf(IsNull(rsingaa!importe), 0, rsingaa!importe)) / 1000
 If Not IsNull(rsingaa!importe) And rsingaa!importe <> 0 Then
    seaa = (((IIf(IsNull(rsing!importe), 0, rsing!importe)) / IIf(IsNull(rsingaa!importe), 0, rsingaa!importe)) * 100) - 100
Else
    seaa = 0
End If
 
 
'----------------------------------------------------------

fechaini = "01/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")

If Product = 1 Then
   sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
  sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(BON_FAC_BONGRA)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa <= " & cboMes.ListIndex + 1 & _
       " and producto = " & Product
       
Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

fechaini = "01/01/" & cboAño.Text
FechaIniAA = "01/01/" & Val(cboAño.Text) - 1

 
 sqls = "select Importe=IsNull(Sum(Case" & _
              " When (CarAbo='C' )" & _
              " Then Importe " & _
              " When (CarAbo='A' ) " & _
              " Then Importe*-1 " & _
              " End),0) " & _
"From clientes_movimientos" & _
 " where fecha between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
 " and tipo_mov in (12,13,88)"

Set rsing = New ADODB.Recordset
rsing.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = "select Importe=IsNull(Sum(Case" & _
              " When (CarAbo='C' )" & _
              " Then Importe " & _
              " When (CarAbo='A' ) " & _
              " Then Importe*-1 " & _
              " End),0) " & _
"From clientes_movimientos" & _
" where fecha between '" & FechaIniAA & "' and '" & FechaFinAA & " 23:59:00'" & _
" and tipo_mov in (12,13,88)"
 
 
Set rsingaa = New ADODB.Recordset
rsingaa.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = "select ingxtarjeta from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & Val(cboAño.Text) - 1 & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


If rsPres.EOF Then
    ventapresupacum = 0
    MsgBox "No esta capturado el presupuesto para el periodo del año anterior"
Else
    ventapresupacum = rsPres!ingxtarjeta
End If


VentaAcum = IIf(IsNull(rsing!importe), 0, rsing!importe) / 1000
PorcSVtaAcum = (IIf(IsNull(rsing!importe), 0, rsing!importe) / IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) * 100
PresupAcum = (IIf(IsNull(rsing!importe), 0, rsing!importe) - ventapresupacum) / 1000
AAntAcum = (IIf(IsNull(rsing!importe), 0, rsing!importe) - IIf(IsNull(rsingaa!importe), 0, rsingaa!importe)) / 1000

If Not IsNull(rsingaa!importe) And rsingaa!importe <> 0 Then
    seaaacum = (((IIf(IsNull(rsing!importe), 0, rsing!importe)) / IIf(IsNull(rsingaa!importe), 0, rsingaa!importe)) * 100) - 100
Else
    seaaacum = 0
End If
 

sqls = " exec sp_EdoResultados  @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
       " , @Concepto = 4" & _
       " , @Mreal = " & Format(Venta, "#########.00") & _
       " , @Porcsvta = " & Format(PorcSVta, "#########.00") & _
       " , @Presup =  " & Format(Presup, "#########.00") & _
       " , @aant = " & Format(aant, "#########.00") & _
       " , @seaa = " & Format(seaa, "#########.00") & _
       " , @acumreal = " & Format(VentaAcum, "#########.00") & _
       " , @acumporcsvta = " & Format(PorcSVtaAcum, "#########.00") & _
       " , @acumpresup = " & Format(PresupAcum, "#########.00") & _
       " , @acumaant= " & Format(AAntAcum, "#########.00") & _
       " , @acumseaa = " & Format(seaaacum, "########.00")

cnxBD.Execute sqls, intRegistros

End Sub
Sub CalculaIngComProv()
Dim rsvta As ADODB.Recordset, rsvtaAA As ADODB.Recordset, rsPres As ADODB.Recordset


fechaini = Format(cboMes.ListIndex + 1, "00") + "/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")

FechaIniAA = Format(cboMes.ListIndex + 1, "00") + "/01/" + CStr(Val(cboAño.Text) - 1)
FechaFinAA = FechaFinMes(cboMes.ItemData(cboMes.ListIndex), CStr(Val(cboAño.Text) - 1))
      

If Product = 1 Then
   sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
    sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa = " & cboMes.ListIndex + 1 & _
       " and producto = " & Product
       
Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

 
sqls = "select Reembolsos from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    IngPresup = 0
    MsgBox "No esta capturado el presupuesto para este periodo", vbInformation, "Falta capturar presupuesto"
Else
    IngPresup = rsPres!reembolsos
End If


sqls = " select sum(comision + ivacomision) Reembolsos from facturascom" & _
       " where fecha between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and status<=2"
Set rsRem = New ADODB.Recordset
rsRem.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = " select sum(comision + ivacomision) Reembolsos from facturascom" & _
       " where fecha between '" & FechaIniAA & "' and '" & FechaFinAA & " 23:59:00'" & _
       " and status<=2"
Set rsremaa = New ADODB.Recordset
rsremaa.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

Venta = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos)) / 1000
PorcSVta = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) / IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) * 100
ant = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) - IIf(IsNull(rsremaa!reembolsos), 0, rsremaa!reembolsos)) / 1000
If rsPres.EOF Then
    Presup = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos)) / 1000
Else
    Presup = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) - IIf(IsNull(rsPres!reembolsos), 0, rsPres!reembolsos)) / 1000
End If


If Not (IsNull(rsremaa!reembolsos)) Then
    seaa = (((IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos)) / IIf(IsNull(rsremaa!reembolsos), 0, rsremaa!reembolsos)) * 100) - 100
Else
    seaa = 0
End If
 
 
fechaini = "01/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")


If Product = 1 Then
    sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas " & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
    sqls = " select isnull(sum(BON_FAC_BONGRA),0) Ventas " & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(valor+ivavalor) Ventas from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa <= " & cboMes.ListIndex + 1 & _
       " and producto = " & Product
       
Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

fechaini = "01/01/" & cboAño.Text
FechaIniAA = "01/01/" & Val(cboAño.Text) - 1

sqls = " select sum(comision + ivacomision) Reembolsos from facturascom" & _
       " where fecha between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and status<=2"
Set rsRem = New ADODB.Recordset
rsRem.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(comision + ivacomision) Reembolsos from facturascom" & _
       " where fecha between '" & FechaIniAA & "' and '" & FechaFinAA & " 23:59:00'" & _
       " and status<=2"
Set rsremaa = New ADODB.Recordset
rsremaa.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = "select * from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo between  " & Val(cboAño.Text) & "01" & " and " & Val(cboAño.Text) & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


If rsPres.EOF Then
    ventapresupacum = 0
    MsgBox "No esta capturado el presupuesto para el periodo del año anterior"
Else
    ventapresupacum = rsPres!reembolsos
End If


VentaAcum = IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) / 1000
PorcSVtaAcum = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) / IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas)) * 100
PresupAcum = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) - ventapresupacum) / 1000
AAntAcum = (IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) - IIf(IsNull(rsremaa!reembolsos), 0, rsremaa!reembolsos)) / 1000
If Not IsNull(rsremaa!reembolsos) Then
    seaaacum = (((IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos)) / IIf(IsNull(rsremaa!reembolsos), 0, rsremaa!reembolsos)) * 100) - 100
Else
    seaaacum = 0
End If
aant = IIf(IsNull(aant), 0, aant)

sqls = " exec sp_EdoResultados  @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
       " , @Concepto = 3" & _
       " , @Mreal = " & Format(Venta, "#########.00") & _
       " , @Porcsvta = " & Format(PorcSVta, "#########.00") & _
       " , @Presup =  " & Format(Presup, "#########.00") & _
       " , @aant = " & Format(aant, "#########.00") & _
       " , @seaa = " & Format(seaa, "#########.00") & _
       " , @acumreal = " & Format(VentaAcum, "#########.00") & _
       " , @acumporcsvta = " & Format(PorcSVtaAcum, "#########.00") & _
       " , @acumpresup = " & Format(PresupAcum, "#########.00") & _
       " , @acumaant= " & Format(AAntAcum, "#########.00") & _
       " , @acumseaa = " & Format(seaaacum, "########.00")

cnxBD.Execute sqls, intRegistros
End Sub

Sub CalculaIngcomClientes()
Dim rsvta As ADODB.Recordset, rsvtaAA As ADODB.Recordset, rsPres As ADODB.Recordset

fechaini = Format(cboMes.ListIndex + 1, "00") + "/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")

FechaIniAA = Format(cboMes.ListIndex + 1, "00") + "/01/" + CStr(Val(cboAño.Text) - 1)
FechaFinAA = FechaFinMes(cboMes.ItemData(cboMes.ListIndex), CStr(Val(cboAño.Text) - 1))


If Product = 1 Then
   sqls = " select  isnull(sum(bon_fac_bonexe),0) Ventas ,isnull(sum(bon_fac_comision),0)  IngComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
   sqls = " select  isnull(sum(BON_FAC_BONGRA),0) Ventas ,isnull(sum(bon_fac_comision),0)  IngComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(valor+ivavalor) Ventas , sum(comision) IngComCte from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa = " & cboMes.ListIndex + 1 & _
       " and producto = " & Product
       
Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

 
sqls = "select IngComCte  from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    IngPresup = 0
    MsgBox "No esta capturado el presupuesto para este periodo", vbInformation, "Falta capturar presupuesto"
    Presup = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte)) / 1000
Else
    IngPresup = rsPres!IngComCte
    Presup = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) - IIf(IsNull(rsPres!IngComCte), 0, rsPres!IngComCte)) / 1000
End If


Venta = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte)) / 1000
PorcSVta = (rsvta!IngComCte / rsvta!Ventas) * 100
aant = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) - IIf(IsNull(rsvtaAA!IngComCte), 0, rsvtaAA!IngComCte)) / 1000

seaa = (((IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte)) / IIf(IsNull(rsvtaAA!IngComCte), 1, rsvtaAA!IngComCte)) * 100) - 100

fechaini = "01/01/" + cboAño.Text
fechafin = Format(FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text), "mm/dd/yyyy")
 
If Product = 1 Then
   sqls = " select  isnull(sum(bon_fac_bonexe),0) Ventas ,isnull(sum(bon_fac_comision),0)  IngComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
Else
   sqls = " select  isnull(sum(BON_FAC_BONGRA),0) Ventas ,isnull(sum(bon_fac_comision),0)  IngComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & " 23:59:00'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"
End If

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(valor+ivavalor) Ventas, sum(comision) IngComCte from ventaxmesxcliente " & _
       " Where anoventa = " & Val(cboAño.Text) - 1 & _
       " and mesventa <= " & cboMes.ListIndex + 1 & _
       " and producto = " & Product

Set rsvtaAA = New ADODB.Recordset
rsvtaAA.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = "select IngComCte from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo between  " & Val(cboAño.Text) & "01" & " and " & Val(cboAño.Text) & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If rsPres.EOF Then
    ventapresupacum = 0
    MsgBox "No esta capturado el presupuesto para el periodo del año anterior", vbInformation, "Presupuesto..."
Else
    ventapresupacum = rsPres!IngComCte
End If


VentaAcum = IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) / 1000
PorcSVtaAcum = (rsvta!IngComCte / rsvta!Ventas) * 100
PresupAcum = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) - ventapresupacum) / 1000
AAntAcum = (IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) - IIf(IsNull(rsvtaAA!IngComCte), 0, rsvtaAA!IngComCte)) / 1000
seaaacum = (((IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte)) / IIf(IsNull(rsvtaAA!IngComCte), 1, rsvtaAA!IngComCte)) * 100) - 100


sqls = " exec sp_EdoResultados  @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
       " , @Concepto = 2" & _
       " , @Mreal = " & Format(Venta, "#########.00") & _
       " , @Porcsvta = " & Format(PorcSVta, "#########.00") & _
       " , @Presup =  " & Format(Presup, "#########.00") & _
       " , @aant = " & Format(aant, "#########.00") & _
       " , @seaa = " & Format(seaa, "#########.00") & _
       " , @acumreal = " & Format(VentaAcum, "#########.00") & _
       " , @acumporcsvta = " & Format(PorcSVtaAcum, "#########.00") & _
       " , @acumpresup = " & Format(PresupAcum, "#########.00") & _
       " , @acumaant= " & Format(AAntAcum, "#########.00") & _
       " , @acumseaa = " & Format(seaaacum, "########.00")

cnxBD.Execute sqls, intRegistros

End Sub


Private Sub cboAño_Click()
   BuscaDatos
End Sub

Private Sub cboAño_Change()
   BuscaDatos
End Sub

Private Sub cboMes_Click()
   BuscaDatos
End Sub

Private Sub cboMes_Change()
   BuscaDatos
End Sub

Private Sub cmdPresentar_Click()
On Error GoTo err_gral
   If cboMes = "" Then
      MsgBox "Primero seleccione el mes que desea calcular", vbCritical, "Mes..."
      cboMes.SetFocus
      Exit Sub
   End If
   If cboAño = "" Then
      MsgBox "Primero seleccione el año que desea calcular", vbCritical, "Año..."
      cboAño.SetFocus
      Exit Sub
   End If
   CalculaVentas
   CalculaIngcomClientes
   CalculaIngComProv
   CalculaIngxTarjetas

  ' GrabaDatos
   Imprime crptToWindow
   Exit Sub

err_gral:
   MsgBox "Error " & ERR.Number & ":" & ERR.Description, , "Reporte de Presupuestos"
  'If cnxBD.BeginTrans > 0 Then cnxBD.RollbackTrans
   Call doErrorLog(gnBodega, "FACBE", ERR.Number, ERR.Description, Usuario, "frmRepPresup.cmdpresentar")
   Resume Next
   MsgBar "", False
End Sub

Sub Imprime(Destino)
Dim Result As Integer
    MsgBar "Generando Reporte", True
    Limpia_CryReport
    mdiMain.cryReport.Connect = "driver={SQL Server};server=" & gstrServidor & ";uid=Operaciones;pwd=" & gpwdDataBase & ";database=" & gstrDataBase
    mdiMain.cryReport.ReportFileName = gPath & "\Reportes\rptEdoResultados.rpt"
    mdiMain.cryReport.Destination = Destino
    mdiMain.cryReport.StoredProcParam(0) = cboAño & Format(cboMes.ItemData(cboMes.ListIndex), "00")
    
    On Error Resume Next
    Result = mdiMain.cryReport.PrintReport
    MsgBar "", False
    If Result <> 0 Then
        MsgBox "Error: " & mdiMain.cryReport.LastErrorNumber & " " & mdiMain.cryReport.LastErrorString, vbCritical
    End If
End Sub
Sub GrabaDatos()
Dim rsvta As ADODB.Recordset
Dim rsRem As ADODB.Recordset
Dim rsPres As ADODB.Recordset
Dim rsInt As ADODB.Recordset
Dim fechaini As String, fechafin As String

On Error GoTo err_gral

fechaini = Format(cboMes.ListIndex + 1, "00") + "/01/" + cboAño.Text
fechafin = FechaFinMes(cboMes.ItemData(cboMes.ListIndex), cboAño.Text)

CalculaVentas

sqls = " select isnull(sum(bon_fac_bonexe),0) Ventas , isnull(sum(bon_fac_comision),0)  IngComCte,( sum(bon_fac_comision)  /sum(bon_fac_bonexe)) * 100 PorcComCte" & _
       " From bon_factura" & _
       " where  bon_fac_fechaemi between '" & fechaini & "' and '" & fechafin & "'" & _
       " and  bon_fac_tpobon = " & Product & _
       " and bon_fac_status = 1"

Set rsvta = New ADODB.Recordset
rsvta.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly


sqls = " select  isnull(sum(ImporteT),0) Reembolsos, isnull(sum(comision),0) IngComProv,( sum(comision)/sum(ImporteT) )*100 PorcComProv " & _
      " From facturascom " & _
      " where Fecha between '" & fechaini & "' and '" & fechafin & "'" & _
      " and status <=2"

Set rsRem = New ADODB.Recordset
rsRem.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select sum(Rendimiento) Inginv" & _
        " from interesesbe" & _
        " where periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")

Set rsInt = New ADODB.Recordset
rsInt.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

sqls = " select isnull(sum(subtotal),0) IngXTarjetas from fm_facturas" & _
       " Where rubro in (12,13)" & _
       " and fecha between '" & fechaini & "' and  '" & fechafin & "'"
Set rsBD = New ADODB.Recordset
rsBD.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsBD.EOF Then
    txtIngXtarjeta = CDbl(rsBD!IngXTarjetas)
End If

sqls = " exec sp_GrabaDatosPresup @Tipo = 1" & _
       " , @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
       " , @Ventas = " & IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) & _
       " , @Reembolsos = " & IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) & _
       " , @IngComCte = " & IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) & _
       " , @PorcComCte = " & IIf(IsNull(rsvta!PorcComCte), 0, rsvta!PorcComCte) & _
       " , @IngComProv = " & IIf(IsNull(rsRem!IngComProv), 0, rsRem!IngComProv) & _
       " , @PorcComProv = " & IIf(IsNull(rsRem!PorcComProv), 0, rsRem!PorcComProv) & _
       " , @InvProm = " & CDbl(txtInvProm) & _
       " , @TasaProm = " & CDbl(txtTasaProm) & _
       " , @IngXTarjeta = " & CDbl(txtIngXtarjeta) & _
       " , @IngInv = " & IIf(IsNull(rsInt!inginv), 0, rsInt!inginv)
cnxBD.Execute sqls, intRegistros
   
   'presupuestado
sqls = "select * from presupuestomensualbe" & _
       " where tipo = 0" & _
       " and periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00")


Set rsPres = New ADODB.Recordset
rsPres.Open sqls, cnxBD, adOpenForwardOnly, adLockReadOnly

If Not rsPres.EOF Then
   
   sqls = " exec sp_GrabaDatosPresup @Tipo = 2" & _
          " , @Periodo = " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00") & _
          " , @Ventas = " & IIf(IsNull(rsvta!Ventas), 0, rsvta!Ventas) - IIf(IsNull(rsPres!Ventas), 0, rsPres!Ventas) & _
          " , @Reembolsos = " & IIf(IsNull(rsRem!reembolsos), 0, rsRem!reembolsos) - IIf(IsNull(rsPres!reembolsos), 0, rsPres!reembolsos) & _
          " , @IngComCte = " & IIf(IsNull(rsvta!IngComCte), 0, rsvta!IngComCte) - IIf(IsNull(rsPres!IngComCte), 0, rsPres!IngComCte) & _
          " , @PorcComCte = " & IIf(IsNull(rsvta!PorcComCte), 0, rsvta!PorcComCte) - IIf(IsNull(rsPres!PorcComCte), 0, rsPres!PorcComCte) & _
          " , @IngComProv = " & IIf(IsNull(rsRem!IngComProv), 0, rsRem!IngComProv) - IIf(IsNull(rsPres!IngComProv), 0, rsPres!IngComProv) & _
          " , @PorcComProv = " & IIf(IsNull(rsRem!PorcComProv), 0, rsRem!PorcComProv) - IIf(IsNull(rsPres!PorcComProv), 0, rsPres!PorcComProv) & _
          " , @InvProm = " & CDbl(txtInvProm) - CDbl(rsPres!InvProm) & _
          " , @TasaProm = " & CDbl(txtTasaProm) - CDbl(rsPres!TasaProm) & _
          " , @IngXTarjeta = " & CDbl(txtIngXtarjeta) - CDbl(rsPres!ingxtarjeta) & _
          " , @IngInv = " & CDbl(IIf(IsNull(rsInt!inginv), 0, rsInt!inginv)) - IIf(IsNull(rsPres!inginv), 0, rsPres!inginv)
   cnxBD.Execute sqls, intRegistros

Else
    MsgBox "No se ha dado de alta el presupuesto para el periodo " & cboAño.Text & Format(cboMes.ItemData(cboMes.ListIndex), "00"), vbExclamation, "Problemas en  presupuesto"
    Screen.MousePointer = 1
End If
rsvta.Close
Set rsvta = Nothing
rsRem.Close
Set rsRem = Nothing
rsPres.Close
Set rsPres = Nothing
rsBD.Close
Set rsBD = Nothing

Exit Sub

err_gral:
    MsgBox ERR.Number & " " & ERR.Description
    Resume Next
    Exit Sub
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub
Private Sub Form_Activate()
 Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
End Sub

Private Sub Form_Load()
   Dim anom As Integer
   Set mclsAniform = New clsAnimated
   CboPosiciona cboMes, Month(Date)
   
   For anom = Year(Date) To Year(Date) - 6 Step -1
       cboAño.AddItem anom
   Next
   cboAño.Text = Year(Date)
      
   cboProducto.Clear
   Call CargaComboBE(cboProducto, "sp_sel_productobe 'BE','','Cargar'")
   Call LeeproductoBE(cboProducto, "sp_sel_productobe 'BE','" & Trim(cboProducto.Text) & " ','Leer'")
   cboProducto.Text = UCase("Despensa Total")
End Sub
Sub BuscaDatos()
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call mclsAniform.Animated(Me, eHIDE, 250, AW_BLEND)
End Sub
