VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMCON 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20355
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   20355
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "FRMCON.frx":0000
      Left            =   8640
      List            =   "FRMCON.frx":0028
      TabIndex        =   9
      Text            =   "Mes"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "FRMCON.frx":005C
      Left            =   8640
      List            =   "FRMCON.frx":006F
      TabIndex        =   8
      Text            =   "Año"
      Top             =   9120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "FRMCON.frx":0091
      Left            =   8640
      List            =   "FRMCON.frx":00F2
      TabIndex        =   7
      Text            =   "Dia"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "FRMCON.frx":0169
      Left            =   6480
      List            =   "FRMCON.frx":017C
      TabIndex        =   6
      Text            =   "Año"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "FRMCON.frx":019E
      Left            =   6480
      List            =   "FRMCON.frx":01C6
      TabIndex        =   5
      Text            =   "Mes"
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "FRMCON.frx":01FA
      Left            =   6480
      List            =   "FRMCON.frx":025B
      TabIndex        =   4
      Text            =   "Dia"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "FRMCON.frx":02D2
      Left            =   7560
      List            =   "FRMCON.frx":02DF
      TabIndex        =   2
      Text            =   "Opciones"
      Top             =   7080
      Width           =   2055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   4695
      Left            =   8760
      TabIndex        =   1
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4695
      Left            =   480
      TabIndex        =   0
      Top             =   1920
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   12298
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "FACTURAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8400
      TabIndex        =   10
      Top             =   480
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   7080
      X2              =   12480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image13 
      Height          =   1200
      Left            =   5640
      Picture         =   "FRMCON.frx":02FE
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   1860
      Left            =   960
      Picture         =   "FRMCON.frx":0911
      Top             =   0
      Width           =   3825
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   5520
      Picture         =   "FRMCON.frx":2C30
      Top             =   10080
      Width           =   1755
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   10320
      Picture         =   "FRMCON.frx":3685
      Top             =   10080
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   7920
      Picture         =   "FRMCON.frx":56C0
      Top             =   10080
      Width           =   1800
   End
   Begin VB.Label label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   16680
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   16395
      Left            =   120
      Picture         =   "FRMCON.frx":752E
      Top             =   120
      Width           =   17625
   End
End
Attribute VB_Name = "FRMCON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As Integer
Dim s, a As Integer
Private Sub Combo1_Click()
  If Combo1.Text = "Desde" Then
        Combo2.Visible = True
        Combo3.Visible = True
        Combo4.Visible = True
        Combo5.Visible = False
        Combo6.Visible = False
        Combo7.Visible = False
    End If
    If Combo1.Text = "Hasta" Then
        Combo5.Visible = True
        Combo6.Visible = True
        Combo7.Visible = True
        Combo2.Visible = False
        Combo3.Visible = False
        Combo4.Visible = False
    End If
    If Combo1.Text = "Desde/Hasta" Then
        Combo5.Visible = True
        Combo6.Visible = True
        Combo7.Visible = True
        Combo2.Visible = True
        Combo3.Visible = True
        Combo4.Visible = True
    End If
End Sub

Private Sub Command1_Click()
With RSVENTAS_ELIMINADAS
    .AddNew
    !IDVENTAS = DataGrid1.Columns(0).Text
    !fecha = DataGrid1.Columns(1).Text
    !CEDULACLIENTE = DataGrid1.Columns(2).Text
    !CEDULADUENO = DataGrid1.Columns(3).Text
 
    End With
    q = rsFactura.RecordCount
    For X = 1 To q
    With RSFACTURA_ELIMINADAS

    !IDFACTURA = DataGrid2.Columns(0).Text
    !IDPRODUCTO = DataGrid2.Columns(1).Text
    !CANTIDAD = DataGrid2.Columns(2).Text
    !PRECIO = DataGrid2.Columns(3).Text
    !IDVENTAS = DataGrid2.Columns(4).Text
    
    End With
    rsFactura.MoveNext
    Next
    With RSVEN
    .Delete
    .MoveFirst
    End With
    For X = 1 To q
    With rsFactura
 
    .Requery
    .Delete
    .MoveNext
    If .EOF Then Exit Sub
    End With
    Next
    
End Sub


Private Sub Command2_Click()
Dim s, a As String
    With RSVEN
        If .State = 1 Then .Close
        If Combo1.Text = "Desde" Then
            s = "#" & desde.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA)>= " & s & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
        If Combo1.Text = "Hasta" Then
            s = "#" & hasta.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA)<= " & s & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
    End With
Set DataGrid1.DataSource = RSVEN
End Sub

Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub DataGrid1_Click()

  Label1 = DataGrid1.Columns(0).Text
    With rsFactura
        Dim s As String
        s = "%" & Label1.Caption & "%"
        If .State = 1 Then .Close
        .Open "Select * From FACTURA Where [IDVENTAS] Like '" & s & "'"
        Set DataGrid2.DataSource = rsFactura
    End With
End Sub


Private Sub Form_Load()
FACTURA_ELIMINADA
VENTAS_ELIMINADAS
tablaVENTAS

factura
Set DataGrid1.DataSource = RSVEN
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img\bus1.jpg")
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\img\bus1.jpg")

Dim s, a As String
    With RSVEN
If .State = 1 Then .Close
        If Combo1.Text = "Desde" Then
            s = "#" & Combo3.Text & "/" & Combo2.Text & "/" & Combo4.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA) >= " & s & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
        If Combo1.Text = "Hasta" Then
        a = "#" & Combo3.Text & "/" & Combo2.Text & "/" & Combo4.Text & "#"
            .Open "Select * From VENTAS Where ((FECHA)<= " & a & ")", CN, adOpenStatic, adLockBatchOptimistic
        End If
        If Combo1.Text = "Desde/Hasta" Then
            If .State = 1 Then .Close
            s = "#" & desde.Text & "#"
            a = "#" & hasta.Text & "#"
        End If
End With
Set DataGrid1.DataSource = RSVEN
End Sub


Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\img\ELI0.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\img\ELI0.jpg")
With RSVENTAS_ELIMINADAS
    .AddNew
    !IDVENTAS = DataGrid1.Columns(0).Text
    !fecha = DataGrid1.Columns(1).Text
    !CEDULACLIENTE = DataGrid1.Columns(2).Text
    !CEDULADUENO = DataGrid1.Columns(3).Text
 
    End With
    q = rsFactura.RecordCount
    For X = 1 To q
    With RSFACTURA_ELIMINADAS

    !IDFACTURA = DataGrid2.Columns(0).Text
    !IDPRODUCTO = DataGrid2.Columns(1).Text
    !CANTIDAD = DataGrid2.Columns(2).Text
    !PRECIO = DataGrid2.Columns(3).Text
    !IDVENTAS = DataGrid2.Columns(4).Text
    
    End With
    rsFactura.MoveNext
    Next
    With RSVEN
    .Delete
    .MoveFirst
    End With
    For X = 1 To q
    With rsFactura
 
    .Requery
    .Delete
    .MoveNext
    If .EOF Then Exit Sub
    End With
    Next
End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\img0\VEN.jpg")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\img0\VEN.jpg")
Form1.Show
End Sub
