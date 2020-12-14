VERSION 5.00
Begin VB.Form frmRename 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Dateien umbenennen"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRename.frx":0000
   LinkMode        =   1  'Quelle
   LinkTopic       =   "MassRename"
   MaxButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   472
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Vorschau"
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   7080
      Width           =   1815
   End
   Begin VB.ListBox lstPreview 
      Height          =   1410
      Left            =   120
      OLEDropMode     =   1  'Manuell
      TabIndex        =   16
      Top             =   5520
      Width           =   6855
   End
   Begin VB.CommandButton cmdFilenameDelete 
      Caption         =   "&Löschen"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdFilenameEdit 
      Caption         =   "&Bearbeiten..."
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdFilenameDown 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      Picture         =   "frmRename.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   15
      Top             =   2460
      Width           =   375
   End
   Begin VB.CheckBox chkFilename 
      Caption         =   "Datei&namen ändern:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ListBox lstFilename 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      IntegralHeight  =   0   'False
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Width           =   6135
   End
   Begin VB.CommandButton cmdFilenameUp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      Picture         =   "frmRename.frx":0884
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdFilenameAdd 
      Caption         =   "&Hinzufügen..."
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtensionDelete 
      Caption         =   "&Löschen"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtensionEdit 
      Caption         =   "&Bearbeiten..."
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtensionAdd 
      Caption         =   "&Hinzufügen..."
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   4800
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtensionUp 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      Picture         =   "frmRename.frx":0CC6
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   3840
      Width           =   375
   End
   Begin VB.ListBox lstExtension 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      IntegralHeight  =   0   'False
      Left            =   360
      TabIndex        =   3
      Top             =   3840
      Width           =   6135
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Ausführen"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   7800
      Width           =   1815
   End
   Begin VB.CheckBox chkExtension 
      Caption         =   "Datei&endung ändern:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   2055
   End
   Begin VB.ListBox lstFiles 
      Height          =   1410
      Left            =   120
      OLEDropMode     =   1  'Manuell
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton cmdExtensionDown 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6480
      Picture         =   "frmRename.frx":1108
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   4380
      Width           =   375
   End
   Begin VB.Label lblProgressExecute 
      Caption         =   "Fortschritt: 0%"
      Height          =   255
      Left            =   4560
      TabIndex        =   19
      Top             =   7860
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblProgressPreview 
      Caption         =   "Fortschritt: 0%"
      Height          =   255
      Left            =   4560
      TabIndex        =   18
      Top             =   7140
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   8
      X2              =   464
      Y1              =   504
      Y2              =   504
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   464
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ASSOCIATION As String = " => "

Private Sub chkExtension_Click()
    lstExtension.Enabled = chkExtension.Value
    cmdExtensionAdd.Enabled = chkExtension.Value
    cmdExtensionEdit.Enabled = chkExtension.Value
    cmdExtensionDelete.Enabled = chkExtension.Value
    cmdExtensionUp.Enabled = chkExtension.Value
    cmdExtensionDown.Enabled = chkExtension.Value
End Sub

Private Sub chkFilename_Click()
    lstFilename.Enabled = chkFilename.Value
    cmdFilenameAdd.Enabled = chkFilename.Value
    cmdFilenameEdit.Enabled = chkFilename.Value
    cmdFilenameDelete.Enabled = chkFilename.Value
    cmdFilenameUp.Enabled = chkFilename.Value
    cmdFilenameDown.Enabled = chkFilename.Value
End Sub

Private Sub Execute(ByVal bPreview As Boolean)
    Dim i As Integer, j As Integer
    Dim sAbsolute As String, sFilename As String, sPath As String, sExtension As String
    Dim sNewFilename As String, sNewExtension As String
    
    Dim free As Integer
    free = FreeFile
    If bPreview = False Then Open IIf(Right$(App.Path, 1) = "\", App.Path, App.Path & "\") & "rollback.bat" For Output As #free
    
    For i = 0 To lstFiles.ListCount - 1
        On Error GoTo err_Single_File
        
        'Verschiedene Teile des Pfadnamens extrahieren
        sAbsolute = lstFiles.List(i)
        sPath = Mid$(sAbsolute, 1, InStrRev(sAbsolute, "\"))
        sFilename = Mid$(sAbsolute, InStrRev(sAbsolute, "\") + 1)
        If (InStr(1, sFilename, ".") > 0) Then
            sExtension = Mid$(sFilename, InStrRev(sFilename, ".") + 1)
            sFilename = Mid$(sFilename, 1, InStrRev(sFilename, ".") - 1)
        Else
            sExtension = ""
        End If
        
        'Dateinamen ändern
        sNewFilename = sFilename
        If chkFilename.Value = 1 Then
            Dim sCondition As String, sAction As String, pos As Long, pos2 As Long
            For j = 0 To lstFilename.ListCount - 1
                sCondition = Mid$(lstFilename.List(j), 1, InStr(1, lstFilename.List(j), ASSOCIATION) - 1)
                sAction = Mid$(lstFilename.List(j), InStr(1, lstFilename.List(j), ASSOCIATION) + Len(ASSOCIATION))
                
                If MatchesCondition(sFilename, sCondition) Then
                    If InStr(1, sAction, ":") > 0 Then
                        Select Case UCase$(Mid$(sAction, 2, InStr(1, sAction, ":") - 2))
                            Case "ENTFERNE" '"Zeichenfolge",Limit
                                pos = InStr(1, sAction, """")
                                pos2 = InStrRev(sAction, ",")
                                sNewFilename = Replace(sNewFilename, Mid$(sAction, pos + 1, InStrRev(sAction, """") - pos - 1), "", , CLng(Mid$(sAction, pos2 + 1, Len(sAction) - pos2 - 1)))
                            Case "ERSETZE" '"Zeichenfolge","Ersatz"
                                sAction = Mid$(sAction, 10)
                                pos = InStr(2, sAction, """") 'Ende String1
                                pos2 = InStr(pos + 1, sAction, """") 'Anfang String2
                                sNewFilename = Replace(sNewFilename, Mid$(sAction, 2, pos - 2), Mid$(sAction, pos2 + 1, Len(sAction) - pos2 - 2))
                            Case "ERSETZEALLE" '"Zeichen","Ersatz"
                                sAction = Mid$(sAction, InStr(1, sAction, ":") + 1)
                                pos = InStr(2, sAction, """") 'Ende String1
                                pos2 = InStr(pos + 1, sAction, """") 'Anfang String2
                                Dim s As String, k As Integer, sReplace As String
                                sReplace = Mid$(sAction, pos2 + 1, Len(sAction) - pos2 - 2)
                                s = Mid$(sAction, 2, pos - 2)
                                For k = 1 To Len(s)
                                    sNewFilename = Replace(sNewFilename, Mid$(s, k, 1), sReplace)
                                Next k
                            Case "KLEIN"
                                sAction = Mid$(sAction, 8)
                                pos = CLng(Mid$(sAction, 1, InStr(1, sAction, ",") - 1)) 'Startindex
                                pos2 = InStr(1, sAction, ",")
                                pos2 = CLng(Mid$(sAction, pos2 + 1, Len(sAction) - pos2 - 1)) 'Länge
                                sNewFilename = Mid$(sNewFilename, 1, pos - 1) & LCase$(Mid$(sNewFilename, pos, pos2)) & Mid$(sNewFilename, pos + pos2)
                            Case "GROSS", "GROß"
                                sAction = Mid$(sAction, InStr(1, sAction, ":") + 1)
                                pos = CLng(Mid$(sAction, 1, InStr(1, sAction, ",") - 1)) 'Startindex
                                pos2 = InStr(1, sAction, ",")
                                pos2 = CLng(Mid$(sAction, pos2 + 1, Len(sAction) - pos2 - 1)) 'Länge
                                sNewFilename = Mid$(sNewFilename, 1, pos - 1) & UCase$(Mid$(sNewFilename, pos, pos2)) & Mid$(sNewFilename, pos + pos2)
                            Case "SCHREIBE"
                                sAction = Mid$(sAction, InStr(1, sAction, ":") + 1)
                                pos = CLng(Mid$(sAction, 1, InStr(1, sAction, ",") - 1)) 'Startindex
                                pos2 = InStr(1, sAction, """") + 1 'Erstes Zeichen der Zeichenfolge
                                sNewFilename = Mid$(sNewFilename, 1, pos - 1) & Mid$(sAction, pos2, Len(sAction) - pos2 - 1) & Mid$(sNewFilename, pos)
                        End Select
                    Else
                        Select Case UCase$(Mid$(sAction, 2, Len(sAction) - 2))
                            Case "GROSS", "GROß"
                                sNewFilename = UCase$(sNewFilename)
                            Case "KLEIN"
                                sNewFilename = LCase$(sNewFilename)
                        End Select
                    End If
                End If
            Next j
        End If
        
        'Check for extension
        sNewExtension = sExtension
        If chkExtension.Value = 1 Then
            Dim sFrom As String, sTo As String
            For j = 0 To lstExtension.ListCount - 1
                sFrom = Mid$(lstExtension.List(j), 1, InStr(1, lstExtension.List(j), ASSOCIATION) - 1)
                sTo = Mid$(lstExtension.List(j), InStr(1, lstExtension.List(j), ASSOCIATION) + Len(ASSOCIATION))
                
                If MatchesExtension(sExtension, sFrom) = True Then
                    Select Case sTo
                        Case "<beibehalten>"
                            sNewExtension = sNewExtension
                        Case "<kleinschreiben>"
                            sNewExtension = LCase$(sNewExtension)
                        Case "<großschreiben>"
                            sNewExtension = UCase$(sNewExtension)
                        Case Else
                            sNewExtension = sTo
                    End Select
                End If
            Next j
        Else
            'Alte Erweiterung beibehalten
            sNewExtension = sExtension
        End If
        
        'Vorschau oder Wirklichkeit?
        If bPreview = True Then
            lstPreview.AddItem sPath & sNewFilename & "." & sNewExtension
            If i Mod 4 = 0 Then lblProgressPreview.Caption = "Fortschritt: " & (i / lstFiles.ListCount * 100) & "%"
        Else
            Name sAbsolute As sPath & sNewFilename & "." & sNewExtension
            If i Mod 4 = 0 Then lblProgressExecute.Caption = "Fortschritt: " & (i / lstFiles.ListCount * 100) & "%"
            Print #free, "ren """ & sPath & sNewFilename & "." & sNewExtension & """ """ & sAbsolute & """"
        End If
        GoTo next_file
        
err_Single_File:
        Err.Clear
next_file:
    Next i
    
    If bPreview = False Then Close #free
End Sub

Private Function MatchesCondition(ByVal sFilename As String, ByRef sCondition As String) As Boolean
    If sCondition = "<IMMER>" Then
        MatchesCondition = True
        Exit Function
    End If
    sFilename = UCase$(sFilename)
    sCondition = UCase$(sCondition)
    Select Case Mid$(sCondition, 2, InStr(1, sCondition, ":") - 2)
        Case "ENTHÄLT"
            MatchesCondition = (InStr(1, sFilename, Mid$(sCondition, 11, Len(sCondition) - 11)) > 0)
        Case "BEGINNT"
            MatchesCondition = (InStr(1, sFilename, Mid$(sCondition, 11, Len(sCondition) - 11)) = 1)
        Case "ENDET"
            MatchesCondition = (InStrRev(sFilename, Mid$(sCondition, 8, Len(sCondition) - 8)) = Len(sFilename) - (Len(sCondition) - 9))
        Case "GLEICH"
            MatchesCondition = (sFilename = Mid$(sCondition, 9, Len(sCondition) - 9))
        Case "ÄHNELT"
            MatchesCondition = (sFilename Like Mid$(sCondition, 9, Len(sCondition) - 9))
    End Select
End Function

Private Function MatchesExtension(ByVal sExtension As String, ByRef matchArray As String) As Boolean
    sExtension = LCase$(sExtension)
    matchArray = LCase$(matchArray)
    If sExtension Like matchArray Then
        MatchesExtension = True
        Exit Function
    End If
    If InStr(1, matchArray, sExtension) = 0 Then
        MatchesExtension = False
        Exit Function
    Else
        Dim s() As String, i As Integer
        s = Split(matchArray, "|")
        For i = LBound(s) To UBound(s)
            If sExtension Like s(i) Then
                MatchesExtension = True
                Erase s
                Exit Function
            End If
        Next i
    End If
End Function

Private Sub cmdExecute_Click()
    If lstPreview.ListCount = 0 Then
        If MsgBox("Möchten Sie vorher nicht doch eine Vorschau berechnen?", vbYesNo) = vbYes Then
            cmdPreview_Click
            Exit Sub
        End If
    End If
    
    lblProgressExecute.Visible = True
    Execute False
    lblProgressExecute.Visible = False
    MsgBox "Vorgang abgeschlossen.", vbInformation
End Sub

Private Sub cmdExtensionAdd_Click()
    frmExtension.Show vbModal, Me
    If frmExtension.WasCanceled = False Then
        lstExtension.AddItem frmExtension.GetOldExtension & ASSOCIATION & frmExtension.GetNewExtension
        chkExtension.Value = 1
    End If
    Unload frmExtension
End Sub

Private Sub cmdExtensionDelete_Click()
    If lstExtension.ListIndex <> -1 Then lstExtension.RemoveItem lstExtension.ListIndex
    If lstExtension.ListCount > 0 Then chkExtension.Value = 1
End Sub

Private Sub cmdExtensionDown_Click()
    chkExtension.Value = 1
    If lstExtension.ListIndex > -1 And lstExtension.ListIndex < lstExtension.ListCount - 1 Then
        Dim sTemp As String
        sTemp = lstExtension.List(lstExtension.ListIndex + 1)
        lstExtension.List(lstExtension.ListIndex + 1) = lstExtension.List(lstExtension.ListIndex)
        lstExtension.List(lstExtension.ListIndex) = sTemp
        lstExtension.ListIndex = lstExtension.ListIndex + 1
    End If
End Sub

Private Sub cmdExtensionEdit_Click()
    If lstExtension.ListIndex = -1 Then Exit Sub
    Load frmExtension
    Dim s() As String
    s = Split(lstExtension.Text, ASSOCIATION)
    frmExtension.cmbOldExtension.Text = s(0)
    frmExtension.cmbNewExtension.Text = s(1)
    frmExtension.Show vbModal, Me
    If frmExtension.WasCanceled = False Then
        lstExtension.List(lstExtension.ListIndex) = frmExtension.GetOldExtension & ASSOCIATION & frmExtension.GetNewExtension
        chkExtension.Value = 1
    End If
    Unload frmExtension
End Sub

Private Sub cmdExtensionUp_Click()
    chkExtension.Value = 1
    If lstExtension.ListIndex > 0 Then
        Dim sTemp As String
        sTemp = lstExtension.List(lstExtension.ListIndex - 1)
        lstExtension.List(lstExtension.ListIndex - 1) = lstExtension.List(lstExtension.ListIndex)
        lstExtension.List(lstExtension.ListIndex) = sTemp
        lstExtension.ListIndex = lstExtension.ListIndex - 1
    End If
End Sub

Private Sub cmdFilenameAdd_Click()
    chkFilename.Value = 1
    frmFilename.Show vbModal, Me
    If frmFilename.WasCanceled = False Then
        lstFilename.AddItem frmFilename.GetCondition & ASSOCIATION & frmFilename.GetAction
    End If
    Unload frmFilename
End Sub

Private Sub cmdFilenameDelete_Click()
    If lstFilename.ListIndex = -1 Then Exit Sub
    lstFilename.RemoveItem lstFilename.ListIndex
End Sub

Private Sub cmdFilenameDown_Click()
    chkFilename.Value = 1
    If lstFilename.ListIndex > -1 And lstFilename.ListIndex < lstFilename.ListCount - 1 Then
        Dim sTemp As String
        sTemp = lstFilename.List(lstFilename.ListIndex + 1)
        lstFilename.List(lstFilename.ListIndex + 1) = lstFilename.List(lstFilename.ListIndex)
        lstFilename.List(lstFilename.ListIndex) = sTemp
        lstFilename.ListIndex = lstFilename.ListIndex + 1
    End If
End Sub

Private Sub cmdFilenameEdit_Click()
    If lstFilename.ListCount = -1 Then Exit Sub
    chkFilename.Value = 1
    Load frmFilename
    Dim s() As String
    s = Split(lstFilename.List(lstFilename.ListIndex), ASSOCIATION, 2)
    frmFilename.cmbCondition.Text = s(0)
    frmFilename.cmbAction.Text = s(1)
    frmFilename.Show vbModal, Me
    If frmFilename.WasCanceled = False Then
        lstFilename.List(lstFilename.ListIndex) = frmFilename.GetCondition & ASSOCIATION & frmFilename.GetAction
    End If
    Unload frmFilename
End Sub

Private Sub cmdFilenameUp_Click()
    chkFilename.Value = 1
    If lstFilename.ListIndex > 0 Then
        Dim sTemp As String
        sTemp = lstFilename.List(lstFilename.ListIndex - 1)
        lstFilename.List(lstFilename.ListIndex - 1) = lstFilename.List(lstFilename.ListIndex)
        lstFilename.List(lstFilename.ListIndex) = sTemp
        lstFilename.ListIndex = lstFilename.ListIndex - 1
    End If
End Sub

Private Sub cmdPreview_Click()
    lstPreview.Clear
    lblProgressPreview.Visible = True
    Execute True
    lblProgressPreview.Visible = False
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If Mid$(CmdStr, 1, 3) = "ADD" Then
        lstFiles.AddItem Mid$(CmdStr, 5)
    End If
End Sub

Private Sub Form_Load()
    chkExtension.Value = 1
    lstExtension.AddItem "*" + ASSOCIATION + "<beibehalten>"
    
    chkFilename.Value = 1
    lstFilename.AddItem "<IMMER>" & ASSOCIATION & "<ERSETZEALLE:""._"","" "">"
End Sub

Private Sub lstExtension_DblClick()
    cmdExtensionEdit_Click
End Sub

Private Sub lstFilename_DblClick()
    cmdFilenameEdit_Click
End Sub

Private Sub lstFiles_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 1 To Data.Files.Count
        lstFiles.AddItem Data.Files.Item(i)
    Next i
    Me.Refresh
End Sub
