VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMAGG 
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   Picture         =   "FRMAGG.frx":0000
   ScaleHeight     =   10500
   ScaleWidth      =   15045
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   120
      Picture         =   "FRMAGG.frx":5E39
      ScaleHeight     =   1755
      ScaleWidth      =   4035
      TabIndex        =   10
      Top             =   120
      Width           =   4095
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
      TabIndex        =   2
      Top             =   3120
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
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   6480
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5106
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
            LCID            =   2058
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
            LCID            =   2058
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
   Begin VB.Image Image6 
      Height          =   750
      Left            =   2400
      Picture         =   "FRMAGG.frx":8158
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
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
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Image Image17 
      Height          =   630
      Left            =   8160
      Picture         =   "FRMAGG.frx":A193
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "AGREGAR NUEVO PRODUCTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   4575
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   120
      Picture         =   "FRMAGG.frx":AC25
      Top             =   5400
      Width           =   1800
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
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "IDPRODUCTO"
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
      TabIndex        =   4
      Top             =   4440
      Width           =   2055
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
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2280
      Width           =   4215
   End
End
Attribute VB_Name = "FRMAGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid1_Click()
Label1.Caption = RSINV2!IDPRODUCTO
TXTNUMP.Text = RSINV2!NOMBRE
TXTCOS.Text = RSINV2!CANTIDAD



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

Private Sub Image3_Click()
If Label1.Caption = "" Then Exit Sub
 With RSINV2
 .Find "IDPRODUCTO='" & Label1.Caption & "'"
 !CANTIDAD = Val(!CANTIDAD) + 1
 TXTCOS.Text = !CANTIDAD
 
 .UpdateBatch
 End With
 Set DataGrid1.DataSource = RSINV2
  
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\img\agr1.jpg")
    
    
End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3.Picture = LoadPicture(App.Path & "\img\agr0.jpg")

 
 
 
End Sub


Private Sub Image6_Click()
If Label1.Caption = "" Then Exit Sub
 With RSINV2
 If (!CANTIDAD = 0) Then Exit Sub
 
 .Find "IDPRODUCTO='" & Label1.Caption & "'"
 !CANTIDAD = Val(!CANTIDAD) - 1
  TXTCOS.Text = !CANTIDAD
 .UpdateBatch
 End With
 Set DataGrid1.DataSource = RSINV2
End Sub
Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli1.jpg")
End Sub

Private Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image6.Picture = LoadPicture(App.Path & "\img\eli0.jpg")
    
    End Sub
    
