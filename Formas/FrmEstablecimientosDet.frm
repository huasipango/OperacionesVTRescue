VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmEstablecimientosDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Sucursales"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   615
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Búsqueda"
      TabPicture(0)   =   "FrmEstablecimientosDet.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ImgMenu"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraBusqueda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Captura"
      TabPicture(1)   =   "FrmEstablecimientosDet.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FraCaptura"
      Tab(1).ControlCount=   1
      Begin VB.Frame FraBusqueda 
         Height          =   4800
         Left            =   105
         TabIndex        =   24
         Top             =   345
         Width           =   8415
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3750
            Left            =   90
            TabIndex        =   28
            Top             =   945
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   6615
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.Line Line2 
            X1              =   15
            X2              =   8370
            Y1              =   765
            Y2              =   765
         End
         Begin VB.Label LblEstableB 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2475
            TabIndex        =   27
            Top             =   465
            Width           =   5850
         End
         Begin VB.Label LblNEstableB 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1230
            TabIndex        =   26
            Top             =   465
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Pertenece a:"
            Height          =   240
            Left            =   165
            TabIndex        =   25
            Top             =   465
            Width           =   960
         End
      End
      Begin VB.Frame FraCaptura 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   12
         Top             =   345
         Width           =   8385
         Begin VB.TextBox TxtCP 
            Height          =   315
            Left            =   4335
            TabIndex        =   7
            Top             =   2970
            Width           =   2670
         End
         Begin VB.TextBox TxtColonia 
            Height          =   315
            Left            =   990
            TabIndex        =   6
            Top             =   2970
            Width           =   2670
         End
         Begin VB.TextBox TxtEstacion 
            Height          =   315
            Left            =   990
            MaxLength       =   15
            TabIndex        =   1
            Top             =   930
            Width           =   1380
         End
         Begin VB.TextBox TxtNombre 
            Height          =   315
            Left            =   990
            TabIndex        =   2
            Top             =   1320
            Width           =   6030
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1740
            Width           =   6030
         End
         Begin VB.ComboBox CboCiudad 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2145
            Width           =   6030
         End
         Begin VB.TextBox TxtDomicilio 
            Height          =   315
            Left            =   990
            TabIndex        =   5
            Top             =   2550
            Width           =   6030
         End
         Begin VB.TextBox TxtTel 
            Height          =   315
            Left            =   990
            TabIndex        =   8
            Top             =   3405
            Width           =   2670
         End
         Begin VB.TextBox TxtRFC 
            Height          =   315
            Left            =   990
            TabIndex        =   9
            Top             =   3855
            Width           =   2670
         End
         Begin VB.TextBox TxtContacto 
            Height          =   315
            Left            =   990
            TabIndex        =   10
            Top             =   4290
            Width           =   6030
         End
         Begin VB.Label Label12 
            Caption         =   "C. P."
            Height          =   180
            Left            =   3900
            TabIndex        =   30
            Top             =   3075
            Width           =   360
         End
         Begin VB.Label Label4 
            Caption         =   "Colonia"
            Height          =   180
            Left            =   165
            TabIndex        =   29
            Top             =   3075
            Width           =   675
         End
         Begin VB.Label Label3 
            Caption         =   "Clave"
            Height          =   195
            Left            =   165
            TabIndex        =   23
            Top             =   1020
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Pertenecen a:"
            Height          =   210
            Left            =   150
            TabIndex        =   22
            Top             =   435
            Width           =   1095
         End
         Begin VB.Label LblNEstable 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1275
            TabIndex        =   21
            Top             =   435
            Width           =   855
         End
         Begin VB.Label LblEstable 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2160
            TabIndex        =   20
            Top             =   435
            Width           =   6015
         End
         Begin VB.Line Line1 
            X1              =   30
            X2              =   8355
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Nombre"
            Height          =   210
            Left            =   165
            TabIndex        =   19
            Top             =   1455
            Width           =   915
         End
         Begin VB.Label Label6 
            Caption         =   "Estado"
            Height          =   210
            Left            =   165
            TabIndex        =   18
            Top             =   1845
            Width           =   705
         End
         Begin VB.Label Label7 
            Caption         =   "Ciudad"
            Height          =   210
            Left            =   165
            TabIndex        =   17
            Top             =   2265
            Width           =   600
         End
         Begin VB.Label Label8 
            Caption         =   "Domicilio"
            Height          =   195
            Left            =   165
            TabIndex        =   16
            Top             =   2685
            Width           =   720
         End
         Begin VB.Label Label9 
            Caption         =   "Teléfono"
            Height          =   210
            Left            =   165
            TabIndex        =   15
            Top             =   3540
            Width           =   690
         End
         Begin VB.Label Label10 
            Caption         =   "R.F.C."
            Height          =   240
            Left            =   165
            TabIndex        =   14
            Top             =   3960
            Width           =   585
         End
         Begin VB.Label Label11 
            Caption         =   "Contacto"
            Height          =   270
            Left            =   165
            TabIndex        =   13
            Top             =   4395
            Width           =   705
         End
      End
      Begin MSComctlLib.ImageList ImgMenu 
         Left            =   240
         Top             =   555
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":0038
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":0D12
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":19EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":26C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":33A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEstablecimientosDet.frx":407A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar BarMenu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nueva Captura"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Editar"
            Object.ToolTipText     =   "Modificación"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar Captura"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Limpiar"
            Object.ToolTipText     =   "Limpiar Captura"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmEstablecimientosDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Opcion As Integer
Private Sub BarMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case gNUEVO: Call Nuevo
    Case gEDITAR: Call Editar
    Case gGUARDAR: Call Guardar
    Case gCANCELAR: Call Inicio
    Case gSALIR: Unload Me
End Select
End Sub
Private Sub CboCiudad_GotFocus()
CboCiudad.Clear
If CboEstado.ListIndex <> -1 Then
    Call Llena_Poblacion(CboCiudad, CboEstado.ItemData(CboEstado.ListIndex))
Else
    Call Mensajes(5)
    CboEstado.SetFocus
End If
End Sub

Private Sub CboCiudad_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    TxtDomicilio.SetFocus
End If
End Sub

Private Sub CboEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If CboEstado.ListIndex <> -1 Then
        CboCiudad.SetFocus
    Else
        Call Mensajes(5)
        CboEstado.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
Call CentraFormaMDI(Me)
Call Llena_Estados(CboEstado)
Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gNUEVO)
Call Inicio

LblNEstableB.Caption = Trim$(FrmCatEEstablecimientos.LblFolio.Caption)
LblEstableB.Caption = Trim$(FrmCatEEstablecimientos.TxtNombre.Text)

sql = "Exec Sp_EstablecimientosDet_Sel " & Val(LblNEstableB.Caption) & ""
consulta.Open sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
With consulta
    While Not .EOF
        Grid.AddItem Trim$(!Servicio) & Chr(9) & Trim$(!Nombre)
        .MoveNext
    Wend
End With

If consulta.State = Activo Then
    consulta.Close
End If

End Sub

Private Sub Grid_Click()
Dim Ciudad As Double
If Grid.Rows > 1 Then
    sql = "Exec Sp_EstablecimientosDet_Busqueda " & CDbl(LblNEstableB.Caption) & ",'" & Grid.TextMatrix(Grid.Row, 0) & "'"
    consulta.Open sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    If Not consulta.EOF Then
        TxtEstacion.Text = Trim$(consulta!Servicio)
        TxtNombre.Text = Trim$(consulta!Nombre)
        Call PosicionaComboEnItemData(CboEstado, consulta!estado)
        Ciudad = consulta!Ciudad
        TxtDomicilio.Text = consulta!Domicilio
        TxtTel.Text = consulta!Telefono
        TxtRFC.Text = consulta!Rfc
        TxtContacto.Text = consulta!Contacto
        TxtColonia.Text = consulta!Colonia
        If consulta!CodigoPostal <> 0 Then
            TxtCP.Text = consulta!CodigoPostal
        Else
            TxtCP.Text = vbNullString
        End If
        
        If consulta.State = Activo Then
            consulta.Close
        End If
        
        Call Llena_Poblacion(CboCiudad, CboEstado.ItemData(CboEstado.ListIndex))
        
        Call PosicionaComboEnItemData(CboCiudad, CInt(Ciudad))
        
        LblNEstable.Caption = Trim$(LblNEstableB.Caption)
        LblEstable.Caption = Trim$(LblEstableB.Caption)
        
        SSTab.Tab = 1
        Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gEDITAR)
    End If
    
End If

If consulta.State = Activo Then
    consulta.Close
End If

End Sub

Private Sub TxtColonia_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtCP.SetFocus
End If
End Sub

Private Sub TxtContacto_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    Call Guardar
End If
End Sub

Private Sub TxtCP_KeyPress(KeyAscii As Integer)
If CapNumerica(KeyAscii) = True Then
    If KeyAscii = vbKeyReturn Then
        TxtTel.SetFocus
    End If
Else
    KeyAscii = 0
End If
End Sub

Private Sub TxtDomicilio_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtColonia.SetFocus
End If
End Sub

Private Sub TxtEstacion_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtEstacion.Text <> vbNullString Then
        TxtNombre.SetFocus
    Else
        Call Mensajes(5)
        TxtEstacion.SetFocus
    End If
End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    If TxtNombre.Text <> vbNullString Then
        CboEstado.SetFocus
    Else
        Call Mensajes(5)
        TxtNombre.SetFocus
    End If
End If
End Sub
Private Sub TxtRFC_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtContacto.SetFocus
End If
End Sub

Private Sub TxtTel_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = vbKeyReturn Then
    TxtRFC.SetFocus
End If
End Sub

Private Sub Inicio()
Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gNUEVO)
Grid.Cols = 2
Grid.Rows = 1
Grid.FormatString = "# Estacion|Nombre"
Grid.ColWidth(0) = 2000
Grid.ColWidth(1) = 5700

TxtEstacion.Enabled = True
TxtEstacion.Text = vbNullString
TxtNombre.Text = vbNullString
CboEstado.ListIndex = -1
CboCiudad.ListIndex = -1
TxtDomicilio.Text = vbNullString
TxtColonia.Text = vbNullString
TxtTel.Text = vbNullString
TxtRFC.Text = vbNullString
TxtContacto.Text = vbNullString
TxtCP.Text = vbNullString
FraCaptura.Enabled = False
FraBusqueda.Enabled = True
SSTab.Tab = 0
Opcion = 0
End Sub

Private Sub Nuevo()
Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gGUARDAR)
FraCaptura.Enabled = True
TxtEstacion.Text = vbNullString
TxtNombre.Text = vbNullString
CboEstado.ListIndex = -1
CboCiudad.ListIndex = -1
TxtDomicilio.Text = vbNullString
TxtColonia.Text = vbNullString
TxtTel.Text = vbNullString
TxtRFC.Text = vbNullString
TxtContacto.Text = vbNullString
TxtCP.Text = vbNullString
FraBusqueda.Enabled = False
SSTab.Tab = 1
Opcion = 0
TxtEstacion.SetFocus
LblNEstable.Caption = Trim$(LblNEstableB.Caption)
LblEstable.Caption = Trim$(LblEstableB.Caption)
End Sub

Private Sub Editar()
Call HabilitarDeshabilitarOpcionesMenu(BarMenu, gGUARDAR)
FraCaptura.Enabled = True
TxtEstacion.Enabled = False
FraBusqueda.Enabled = False
SSTab.Tab = 1
Opcion = 1
TxtNombre.SetFocus
LblNEstable.Caption = Trim$(LblNEstableB.Caption)
LblEstable.Caption = Trim$(LblEstableB.Caption)
End Sub

Private Sub Guardar()
Dim Resp As Integer, CP As Double

On Error GoTo errbonos

If TxtEstacion.Text = vbNullString Then
    Call Mensajes(5)
    TxtEstacion.SetFocus
    Exit Sub
End If

If TxtNombre.Text = vbNullString Then
    Call Mensajes(5)
    TxtNombre.SetFocus
    Exit Sub
End If

If CboEstado.ListIndex = -1 Then
    Call Mensajes(5)
    CboEstado.SetFocus
    Exit Sub
End If

If CboCiudad.ListIndex = -1 Then
    Call Mensajes(5)
    CboCiudad.SetFocus
    Exit Sub
End If

If TxtCP.Text = vbNullString Then
    CP = 0
Else
    CP = TxtCP.Text
End If

If Opcion = 0 Then
    Call Mensajes(0)
    If RespMsg = vbYes Then
        sql = "Exec Sp_EstablecimientosDet_Ins " & CDbl(LblNEstable.Caption) & ",'" & Trim$(TxtEstacion.Text) & "','" & Trim$(TxtNombre.Text) & "'," & CboEstado.ItemData(CboEstado.ListIndex) & "," & _
            CboCiudad.ItemData(CboCiudad.ListIndex) & ",'" & Trim$(TxtDomicilio.Text) & "','" & Trim$(TxtTel.Text) & "','" & Trim$(TxtRFC.Text) & "','" & Trim$(TxtContacto.Text) & "','" & Trim$(TxtColonia.Text) & "'," & TxtCP.Text & ""
        consulta.Open sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
Else
    Call Mensajes(2)
    If RespMsg = vbYes Then
        sql = "Exec Sp_EstablecimientosDet_Upd " & CDbl(LblNEstable.Caption) & ",'" & Trim$(TxtEstacion.Text) & "','" & Trim$(TxtNombre.Text) & "'," & CboEstado.ItemData(CboEstado.ListIndex) & "," & _
            CboCiudad.ItemData(CboCiudad.ListIndex) & ",'" & Trim$(TxtDomicilio.Text) & "','" & Trim$(TxtTel.Text) & "','" & Trim$(TxtRFC.Text) & "','" & Trim$(TxtContacto.Text) & "','" & Trim$(TxtColonia.Text) & "'," & TxtCP.Text & ""
        consulta.Open sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    End If
End If

If consulta.State = Activo Then
    consulta.Close
End If


Resp = MsgBox("¿Desea capturar otra sucursal para el mismo grupo?", vbYesNo + vbQuestion, "Servi-Bonos")
If Resp = vbYes Then
    TxtEstacion.Enabled = True
    TxtEstacion.Text = vbNullString
    TxtNombre.Text = vbNullString
    CboEstado.ListIndex = -1
    CboCiudad.ListIndex = -1
    TxtDomicilio.Text = vbNullString
    TxtColonia.Text = vbNullString
    TxtTel.Text = vbNullString
    TxtRFC.Text = vbNullString
    TxtContacto.Text = vbNullString
    TxtCP.Text = vbNullString
    TxtEstacion.SetFocus
    Opcion = 0
Else
    Call Inicio
    
    sql = "Exec Sp_EstablecimientosDet_Sel " & Val(LblNEstableB.Caption) & ""
    consulta.Open sql, Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    With consulta
        While Not .EOF
            Grid.AddItem Trim$(!Servicio) & Chr(9) & Trim$(!Nombre)
            .MoveNext
        Wend
    End With

    If consulta.State = Activo Then
        consulta.Close
    End If
    
End If

Exit Sub

errbonos:
Call Mensajes(6)
Call doErrorLog(gBodega, "SER", ERR.Number, ERR.Description, gUsuario, "modEstablecimientosDet.Guardar")
End Sub
