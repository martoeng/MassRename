VERSION 5.00
Begin VB.Form frmFilename 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Dateiname bearbeiten"
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
   Icon            =   "frmFilename.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   653
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   2453
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.ComboBox cmbAction 
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
      ItemData        =   "frmFilename.frx":000C
      Left            =   360
      List            =   "frmFilename.frx":0028
      TabIndex        =   2
      Text            =   "<ERSETZE:""."","" "">"
      Top             =   1080
      Width           =   4095
   End
   Begin VB.ComboBox cmbCondition 
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
      ItemData        =   "frmFilename.frx":0100
      Left            =   360
      List            =   "frmFilename.frx":0116
      TabIndex        =   0
      Text            =   "<IMMER>"
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label lblAction 
      Caption         =   "Aktion:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label lblCondition 
      Caption         =   "Vorbedingung:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmFilename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Canceled As Boolean
Private sCondition As String
Private sAction As String

Private Sub cmdCancel_Click()
    Canceled = True
    sCondition = ""
    sAction = ""
    Me.Visible = False
End Sub

Private Sub cmdOK_Click()
    Canceled = False
    sCondition = cmbCondition.Text
    sAction = cmbAction.Text
    Me.Visible = False
End Sub

Public Function WasCanceled() As Boolean
    WasCanceled = Canceled
End Function

Public Function GetCondition() As String
    GetCondition = sCondition
End Function

Public Function GetAction() As String
    GetAction = sAction
End Function
