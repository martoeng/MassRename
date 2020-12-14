VERSION 5.00
Begin VB.Form frmExtension 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dateiendung bearbeiten"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExtension.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   2453
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   653
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cmbNewExtension 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmExtension.frx":000C
      Left            =   360
      List            =   "frmExtension.frx":002E
      TabIndex        =   3
      Text            =   "<beibehalten>"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.ComboBox cmbOldExtension 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmExtension.frx":0088
      Left            =   360
      List            =   "frmExtension.frx":009B
      TabIndex        =   1
      Text            =   "*"
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblNewExtension 
      Caption         =   "Neue Dateinamenerweiterung:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblOldExtension 
      Caption         =   "Aktuelle Dateinamenerweiterung:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Canceled As Boolean
Private OldExtension As String
Private NewExtension As String

Private Sub cmdCancel_Click()
    Canceled = True
    OldExtension = ""
    NewExtension = ""
    Me.Visible = False
End Sub

Private Sub cmdOK_Click()
    OldExtension = cmbOldExtension.Text
    NewExtension = cmbNewExtension.Text
    Canceled = False
    Me.Visible = False
End Sub

Public Function WasCanceled() As Boolean
    WasCanceled = Canceled
End Function

Public Function GetOldExtension() As String
    GetOldExtension = OldExtension
End Function

Public Function GetNewExtension() As String
    GetNewExtension = NewExtension
End Function

Private Sub Form_Load()
    cmbOldExtension.AddItem "*", 0
End Sub
