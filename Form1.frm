VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMINV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INVENTARIO - Bazar Jessica"
   ClientHeight    =   12525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20250
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12525
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   3855
      Left            =   10920
      TabIndex        =   19
      Top             =   5160
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6800
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8880
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   8760
      List            =   "Form1.frx":0010
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Frame FR2 
      Caption         =   "Selccione para ordenar: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   13320
      TabIndex        =   14
      Top             =   3240
      Width           =   2175
      Begin VB.OptionButton OP3 
         Caption         =   "NOMBRE"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Op4 
         Caption         =   "CANTIDAD"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame FR 
      Caption         =   "Selección: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   11400
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
      Begin VB.OptionButton Op1 
         Caption         =   "Mayor "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Op2 
         Caption         =   "Menor"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox TXTNUMP 
      DataField       =   "NOMBRE"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox TXTCAN 
      DataField       =   "CANTIDAD"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox TXTCOS 
      DataField       =   "PRECIO"
      DataSource      =   "ADODCINV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":003D
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   5160
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
   Begin MSAdodcLib.Adodc ADODCINV 
      Height          =   615
      Left            =   11520
      Top             =   12000
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAOLA\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\PAOLA\Desktop\PROYECTOFINAL\DATA\BASEINV.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "INVENTARIO"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image11 
      Height          =   750
      Left            =   5880
      Picture         =   "Form1.frx":0054
      Top             =   11520
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image20 
      Height          =   735
      Left            =   14040
      Picture         =   "Form1.frx":23B0
      Top             =   2160
      Width           =   1830
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCAR POR:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   13
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "!!No existe Stock ¡¡"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image19 
      Height          =   1185
      Left            =   0
      Picture         =   "Form1.frx":490C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1200
   End
   Begin VB.Image ImageM 
      Height          =   675
      Left            =   2160
      Picture         =   "Form1.frx":61B4
      Top             =   11640
      Width           =   1725
   End
   Begin VB.Image Image18 
      Height          =   690
      Left            =   3960
      Picture         =   "Form1.frx":6A2A
      Stretch         =   -1  'True
      Top             =   11520
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Image Image17 
      Height          =   630
      Left            =   9960
      Picture         =   "Form1.frx":84F5
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image16 
      Height          =   690
      Left            =   7320
      Picture         =   "Form1.frx":8F87
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image15 
      Height          =   690
      Left            =   10080
      Picture         =   "Form1.frx":9F9D
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image14 
      Height          =   690
      Left            =   9120
      Picture         =   "Form1.frx":AFBB
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image13 
      Height          =   690
      Left            =   8160
      Picture         =   "Form1.frx":BD3B
      Top             =   11760
      Width           =   780
   End
   Begin VB.Image Image12 
      Height          =   1185
      Left            =   9000
      Picture         =   "Form1.frx":CAB3
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Image Image10 
      Height          =   750
      Left            =   240
      Picture         =   "Form1.frx":E35B
      Top             =   11640
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   120
      Top             =   11400
      Width           =   10695
   End
   Begin VB.Image Image9 
      Height          =   1575
      Left            =   -120
      Picture         =   "Form1.frx":100F3
      Stretch         =   -1  'True
      Top             =   11280
      Width           =   11145
   End
   Begin VB.Image Image8 
      Height          =   3495
      Left            =   8160
      Picture         =   "Form1.frx":10330
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   225
   End
   Begin VB.Image Image7 
      Height          =   750
      Left            =   8760
      Picture         =   "Form1.frx":1056D
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image6 
      Height          =   750
      Left            =   5880
      Picture         =   "Form1.frx":123DB
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image5 
      Height          =   750
      Left            =   3960
      Picture         =   "Form1.frx":14416
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image4 
      Height          =   750
      Left            =   2040
      Picture         =   "Form1.frx":16093
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   120
      Picture         =   "Form1.frx":180D0
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   4680
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CANTIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "COSTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "DATOS DE LOS PRODUCTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   1920
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "FRMINV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As Integer

Private Sub CMDBUSCAR_Click()
RSINV.Update
DataGrid1.Refresh
RSINV.Find "idproducto=" & Val(TXTIDPRO.Text)

End Sub

Private Sub CMDMOD_Click()
    MsgBox "Llenar todos los campos de datos de los productos y guardar para agregar correctamente al inventario.", vbInformation, "Dialogo"
    RSINV.MoveLast
    RSINV.AddNew
    RSINV("NOMBRE") = TXTNUMP.Text
    RSINV("PRECIO") = TXTCOS.Text
    RSINV("CANTIDAD") = TXTCAN.Text
    RSINV("IDPROVEEDORES") = 0
    RSINV.Update
    RSINV.Update
    RSINV.MoveLast
    
    
End Sub
'
Private Sub CMDSAL_Click()
If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo) = vbYes Then
            End
    End If
    RSINV.MoveFirst
'BUTONES DE MOVIMIENTO INICIO
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Cantidad" Then
FR.Visible = True
Else
FR.Visible = False
End If
If Combo1.Text <> "" Then
Text1.Enabled = True
End If
Text1.Text = ""
End Sub

Private Sub Command1_Click()
If OP3(0).Value = True Then
RSINV.Sort = "NOMBRE"
ElseIf Op4.Value = True Then
RSINV.Sort = "CANTIDAD"
Else
MsgBox "sleccionar Opción"
End If
Set DataGrid1.DataSource = RSINV



End Sub

Private Sub Command2_Click()

RSINV.MovePrevious

End Sub

Private Sub Command3_Click()

RSINV.MoveNext

End Sub

Private Sub Command4_Click()

RSINV.MoveLast

End Sub
'BUTONES DE MOVIMIENTO END


Private Sub ADODCVENTAS_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub


Private Sub DataGrid1_DblClick()
If DataGrid1.ApproxCount < 1 Then
MsgBox "no ha seleccionado ningun registro", vbExclamation
Exit Sub
Else
    
      TXTIDPRO.Text = DataGrid1.Columns(0).Text
     TXTNUMP.Text = DataGrid1.Columns(1).Text
     TXTCAN.Text = DataGrid1.Columns(2).Text
     TXTCOS.Text = DataGrid1.Columns(3).Text
     
     'TXTIDPROV.Text = DataGrid1.Columns(4).Text
    
    
End If

End Sub

Private Sub DataGrid2_Click()
  C = RSPRO!IDPROVEEDOR
  With RSINV
  If .State = 1 Then .Close
  .Open "select * from INVENTARIO where [IDPROVEEDORES]Like '" & C & "' and CANTIDAD >0 "
  End With
  Set DataGrid1.DataSource = RSINV

Private Sub Form_Load()
    FRMINV.Picture = LoadPicture(App.Path & "\IMG\tst.jpg")
   
    Image2.Picture = LoadPicture(App.Path & "\IMG\logoinv.gif")
    tablaINVENTARIO
    DataGrid1.Columns(1).Width = 5600
    
   
    'Image5.Picture = LoadPicture(App.Path & "\IMG\ed2.jpg")
    b = 0

    If RSINV.State = 1 Then RSINV.Close
    RSINV.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RSINV.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RSINV.Open "select * from  INVENTARIO WHERE ((INVENTARIO.[CANTIDAD])>0)", CN
     Set DataGrid1.DataSource = RSINV
    



End Sub





Private Sub Image1_Click()
If OP3(0).Value = True Then
RSINV.Sort = "NOMBRE"
ElseIf Op4.Value = True Then
RSINV.Sort = "CANTIDAD"
Else
MsgBox "sleccionar Opción"
End If
Set DataGrid1.DataSource = RSINV
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Picture = LoadPicture(App.Path & "\img\O0.jpg")
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image1.Picture = LoadPicture(App.Path & "\img\O1.jpg")
End Sub

Private Sub Image18_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image18.Picture = LoadPicture(App.Path & "\img0\pro1.jpg")
End Sub

Private Sub Image18_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image18.Picture = LoadPicture(App.Path & "\img0\pro0.jpg")
    FRMNUELO.Show
    Unload Me
End Sub



Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\img\agr1.jpg")
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Image3.Picture = LoadPicture(App.Path & "\img\agr0.jpg")
    FRMA.Show

    
   
    Image4.Picture = LoadPicture(App.Path & "\IMG\gua0.jpg")
    'Image5.Picture = LoadPicture(App.Path & "\IMG\ed0.jpg")
    
    
   
    
End Sub


Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4.Picture = LoadPicture(App.Path & "\img\gua1.jpg")
End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    If b = 1 Then
    If TXTNUMP.Text = "" Or TXTCAN.Text = "" Or TXTCOS.Text = "" Then
    Image4.Picture = LoadPicture(App.Path & "\img\gua0.jpg")
    MsgBox "Llenar todos los campos de datos de los productos", vbInformation, "Dialogo"
    Exit Sub
    
    Else
    
    RSINV.Fields("NOMBRE") = TXTNUMP.Text
    RSINV.Fields("CANTIDAD") = TXTCAN.Text
    RSINV.Fields("PRECIO") = TXTCOS.Text
    RSINV.Update
    MsgBox "El registro ha sido actualizado.", vbInformation, "Dialogo"
    b = 0
    DataGrid1.Enabled = True
    Image13.Enabled = True
    Image14.Enabled = True
    Image15.Enabled = True
    Image16.Enabled = True
    Image6.Enabled = True
    Image7.Enabled = True
    End If
    Else
    Image4.Picture = LoadPicture(App.Path & "\IMG\guad2.jpg")
    
    End If
End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image5.Picture = LoadPicture(App.Path & "\img\ed1.jpg")
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image5.Picture = LoadPicture(App.Path & "\img\ed0.jpg")
     If MsgBox("Esta seguro que desea editar un registro?", vbQuestion + vbYesNo) = vbYes Then
     
    DataGrid1.Enabled = False
    Image13.Enabled = False
    Image14.Enabled = False
    Image15.Enabled = False
    Image16.Enabled = False
    Image6.Enabled = False
    Image7.Enabled = False
    Image4.Picture = LoadPicture(App.Path & "\img\gua0.jpg")
    b = 1
     
    End If

End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli1.jpg")
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli0.jpg")
    If MsgBox("Esta seguro que desea eliminar un registro?", vbQuestion + vbYesNo) = vbYes Then
        RSINV.Delete
    End If
End Sub


Private Sub Image7_Click()

tablaINVENTARIO
Label1.Visible = False


X = "%" & Text1.Text & "%"
If Text1.Text = "" Then
    MsgBox "Rellenar la casilla", vbCritical, "Rellenar casilla"
End If
If Combo1.Text = "Proveedor" Then
With RSPRO
If .State = 1 Then .Close
    .Open "select * from PROVEEDORES where [NOMBRE]like '" & X & "'", CN, adOpenStatic, adLockBatchOptimistic
    Set DataGrid2.DataSource = RSPRO
End With
Else
With RSINV
If .State = 1 Then .Close
If Combo1.Text = "Nombre" Then .Open "select * from INVENTARIO where [NOMBRE] like '" & X & "' and CANTIDAD > 0 ", CN, adOpenStatic, adLockBatchOptimistic Else .Open "select * from INVENTARIO where [IDPRODUCTO] like '" & Text1.Text & "' and CANTIDAD > 0", CN, adOpenStatic, adLockBatchOptimistic
If .EOF Or .BOF Then Label1.Visible = True
Set DataGrid1.DataSource = RSINV
End With
End If

If Combo1.Text = "Cantidad" Then
With RSINV
If .State = 1 Then .Close
If Op1.Value = True Then .Open " select * from INVENTARIO where ((CANTIDAD) > " & Text1.Text & ") and CANTIDAD > 0 ", CN, adOpenStatic, adLockBatchOptimistic
If Op2.Value = True Then .Open " select * from INVENTARIO where ((CANTIDAD) < " & Text1.Text & ") and CANTIDAD > 0 ", CN, adOpenStatic, adLockBatchOptimistic
Set DataGrid1.DataSource = RSINV
End With
End If


 'If RSINV.BOF = False And RSIN.EOF = False Then
 'RSINV.Update
    'DataGrid1.Refresh
    


End Sub
Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture(App.Path & "\img\bus1.jpg")
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image7.Picture = LoadPicture(App.Path & "\img\bus0.jpg")
    
   
End Sub


Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.Picture = LoadPicture(App.Path & "\img\ven1.jpg")
End Sub

Private Sub Image10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image10.Picture = LoadPicture(App.Path & "\img\ven0.jpg")
    FRMVENTAS.Show
    Unload Me
End Sub

Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.Picture = LoadPicture(App.Path & "\img\da1.jpg")
End Sub

Private Sub Image11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image11.Picture = LoadPicture(App.Path & "\img\da0.jpg")
    Set RS = CN.Execute("select *from inventario")
    If RS.EOF = False Then
    Set DRINV.DataSource = RS
    DRINV.Show
End If
End Sub
Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri1.jpg")
End Sub

Private Sub Image13_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image13.Picture = LoadPicture(App.Path & "\img\pri0.jpg")
    RSINV.MovePrevious
End Sub

Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig1.jpg")
End Sub

Private Sub Image14_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image14.Picture = LoadPicture(App.Path & "\img\sig0.jpg")
    RSINV.MoveNext
End Sub

Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi1.jpg")
End Sub

Private Sub Image15_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image15.Picture = LoadPicture(App.Path & "\img\fi0.jpg")
    RSINV.MoveLast
End Sub

Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in1.jpg")
End Sub

Private Sub Image16_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image16.Picture = LoadPicture(App.Path & "\img\in0.jpg")
    RSINV.MoveFirst
End Sub
Private Sub Image17_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image17.Picture = LoadPicture(App.Path & "\img\X1.jpg")
End Sub

Private Sub Image17_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Image17.Picture = LoadPicture(App.Path & "\img\X0.jpg")
    If MsgBox("Esta seguro que desea cerrar el formulario?", vbQuestion + vbYesNo, "Inventario") = vbYes Then
        
            End
    End If
    RSINV.MoveFirst
End Sub
Private Sub ImageM_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu1.jpg")
End Sub

Private Sub ImageM_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ImageM.Picture = LoadPicture(App.Path & "\img0\menu0.jpg")
If MsgBox("Esta seguro que desea regresar al menu?", vbQuestion + vbYesNo, "Inventario") = vbYes Then
    FRMMENU.Show
    
    
    Unload Me
    
    End If
End Sub


Sub IDPRODUCTO()

    If RSPRO.State = 1 Then RSPRO.Close
    RSPRO.CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    RSPRO.LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
            
    RSPRO.Open "select * from  INVENTARIO WHERE ((INVENTARIO.[CANTIDAD])>0)", CN
    
End Sub


Private Sub Image9_Click()

FRMAGG.Show


 '.MoveFirst
 With RSINV2
 'For i = 1 To .RecordCount
'If Val(!CANTIDAD) = 0 Then !TEMP = "1" Else !TEMP = "0"
'.UpdateBatch
'.MoveNext
'Next i
 If .State = 1 Then .Close
    .CursorType = adOpenKeyset 'Definimos el tipo de cursor.
    .LockType = adLockOptimistic 'Definimos el tipo de bloqueo.
     a = 1
    .Open "select * from  INVENTARIO WHERE [CANTIDAD] < 1", CN
    
End With
Set FRMAGG.DataGrid1.DataSource = RSINV2

End Sub

Private Sub Image9_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Image9.Picture = LoadPicture(App.Path & "\img\A0.jpg")
End Sub

Private Sub Image9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image9.Picture = LoadPicture(App.Path & "\img\A1..jpg")
End Sub

