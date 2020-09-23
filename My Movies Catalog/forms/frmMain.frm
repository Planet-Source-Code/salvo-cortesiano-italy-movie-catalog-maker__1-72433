VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movies Catalog Maker v1.0.2"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkDownloadMovie 
      Caption         =   "Download Trailer if exist"
      Height          =   255
      Left            =   345
      TabIndex        =   58
      Top             =   6555
      Value           =   1  'Checked
      Width           =   4320
   End
   Begin VB.CheckBox ChkError 
      Caption         =   "On {error} continue parsing"
      Height          =   255
      Left            =   4725
      TabIndex        =   55
      Top             =   6555
      Value           =   1  'Checked
      Width           =   3270
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export Data"
      Enabled         =   0   'False
      Height          =   300
      Left            =   8145
      TabIndex        =   51
      Top             =   6510
      Width           =   1560
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   300
      Left            =   9885
      TabIndex        =   13
      Top             =   6510
      Width           =   1065
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option Search and Title"
      Height          =   705
      Left            =   75
      TabIndex        =   6
      Top             =   6840
      Width           =   11010
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   30
         ScaleHeight     =   450
         ScaleWidth      =   10935
         TabIndex        =   7
         Top             =   210
         Width           =   10935
         Begin VB.CommandButton cmdSearch 
            Caption         =   "Start"
            Height          =   300
            Left            =   9795
            TabIndex        =   12
            Top             =   45
            Width           =   1005
         End
         Begin VB.ComboBox cmbListServer 
            Height          =   330
            ItemData        =   "frmMain.frx":23D2
            Left            =   6165
            List            =   "frmMain.frx":23D9
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   45
            Width           =   3495
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   765
            TabIndex        =   9
            Text            =   "Ricatto D'amore"
            Top             =   75
            Width           =   4155
         End
         Begin VB.Label Label3 
            Caption         =   "Search In:"
            Height          =   240
            Left            =   5040
            TabIndex        =   10
            Top             =   105
            Width           =   1155
         End
         Begin VB.Label Label2 
            Caption         =   "Title:"
            Height          =   225
            Left            =   90
            TabIndex        =   8
            Top             =   105
            Width           =   720
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Movie Info"
      Height          =   5625
      Left            =   3180
      TabIndex        =   4
      Top             =   795
      Width           =   7905
      Begin VB.PictureBox picList 
         BorderStyle     =   0  'None
         Height          =   5325
         Left            =   45
         ScaleHeight     =   5325
         ScaleWidth      =   7815
         TabIndex        =   45
         Top             =   270
         Visible         =   0   'False
         Width           =   7815
         Begin VB.ListBox lstUrl 
            Height          =   4890
            Left            =   60
            TabIndex        =   48
            Top             =   45
            Width           =   7665
         End
         Begin VB.CommandButton cmdSelect 
            Caption         =   "Select"
            Height          =   300
            Left            =   6615
            TabIndex        =   47
            Top             =   4980
            Width           =   1155
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   5160
            TabIndex        =   46
            Top             =   4980
            Width           =   1260
         End
         Begin VB.Label Label4 
            Caption         =   "Select the apropiated Link from the list, then click 'Select' or 'Cancel'..."
            Height          =   405
            Left            =   90
            TabIndex        =   49
            Top             =   4920
            Width           =   4740
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   5400
         Left            =   15
         ScaleHeight     =   5400
         ScaleWidth      =   7830
         TabIndex        =   5
         Top             =   195
         Width           =   7830
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   14
            Left            =   1755
            TabIndex        =   56
            Text            =   "n.a"
            Top             =   3525
            Width           =   6030
         End
         Begin VB.TextBox txtFileds 
            Height          =   1050
            Index           =   13
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Text            =   "frmMain.frx":23F9
            Top             =   4065
            Width           =   7665
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   12
            Left            =   6345
            TabIndex        =   42
            Text            =   "n.a"
            Top             =   3195
            Width           =   1440
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   11
            Left            =   1755
            TabIndex        =   40
            Text            =   "n.a"
            Top             =   3195
            Width           =   3300
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   10
            Left            =   1755
            TabIndex        =   38
            Text            =   "n.a"
            Top             =   2850
            Width           =   6030
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   9
            Left            =   1755
            TabIndex        =   36
            Text            =   "n.a"
            Top             =   2520
            Width           =   6030
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   8
            Left            =   1755
            TabIndex        =   34
            Text            =   "n.a"
            Top             =   2190
            Width           =   6030
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   7
            Left            =   1755
            TabIndex        =   32
            Text            =   "n.a"
            Top             =   1860
            Width           =   6030
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   6
            Left            =   6795
            TabIndex        =   30
            Text            =   "n.a"
            Top             =   1515
            Width           =   990
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   5
            Left            =   1755
            TabIndex        =   28
            Text            =   "n.a"
            Top             =   1515
            Width           =   3900
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   4
            Left            =   6345
            TabIndex        =   26
            Text            =   "n.a"
            Top             =   1170
            Width           =   1425
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   3
            Left            =   1755
            TabIndex        =   24
            Text            =   "n.a"
            Top             =   1170
            Width           =   3900
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   2
            Left            =   1755
            TabIndex        =   20
            Text            =   "n.a"
            Top             =   825
            Width           =   6015
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   1
            Left            =   1755
            TabIndex        =   18
            Text            =   "n.a"
            Top             =   480
            Width           =   6015
         End
         Begin VB.TextBox txtFileds 
            Height          =   285
            Index           =   0
            Left            =   1755
            TabIndex        =   17
            Text            =   "n.a"
            Top             =   120
            Width           =   6015
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Movie Trailer:"
            Height          =   255
            Index           =   14
            Left            =   105
            TabIndex        =   57
            Top             =   3570
            Width           =   1635
         End
         Begin VB.Label lbls 
            Caption         =   "Plot:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Index           =   13
            Left            =   90
            TabIndex        =   44
            Top             =   3840
            Width           =   660
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Date Out:"
            Height          =   255
            Index           =   12
            Left            =   5100
            TabIndex        =   41
            Top             =   3195
            Width           =   1185
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Distribution:"
            Height          =   255
            Index           =   11
            Left            =   90
            TabIndex        =   39
            Top             =   3225
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Production:"
            Height          =   255
            Index           =   10
            Left            =   90
            TabIndex        =   37
            Top             =   2880
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Cast:"
            Height          =   255
            Index           =   9
            Left            =   90
            TabIndex        =   35
            Top             =   2580
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Official Site:"
            Height          =   255
            Index           =   8
            Left            =   90
            TabIndex        =   33
            Top             =   2235
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Directed By:"
            Height          =   255
            Index           =   7
            Left            =   90
            TabIndex        =   31
            Top             =   1920
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Duration:"
            Height          =   255
            Index           =   6
            Left            =   5700
            TabIndex        =   29
            Top             =   1530
            Width           =   1035
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Gender:"
            Height          =   255
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   1560
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Year:"
            Height          =   255
            Index           =   4
            Left            =   5085
            TabIndex        =   25
            Top             =   1185
            Width           =   1215
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Country:"
            Height          =   255
            Index           =   3
            Left            =   90
            TabIndex        =   23
            Top             =   1215
            Width           =   1635
         End
         Begin VB.Label lblStatus 
            Caption         =   "##"
            Height          =   270
            Left            =   120
            TabIndex        =   22
            Top             =   5145
            Width           =   7650
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Original Title:"
            Height          =   255
            Index           =   2
            Left            =   90
            TabIndex        =   21
            Top             =   870
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Link Card:"
            Height          =   255
            Index           =   1
            Left            =   90
            TabIndex        =   19
            Top             =   510
            Width           =   1635
         End
         Begin VB.Label lbls 
            Alignment       =   1  'Right Justify
            Caption         =   "Movie Title:"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   150
            Width           =   1635
         End
      End
   End
   Begin InetCtlsObjects.Inet Inet 
      Left            =   10500
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame frames 
      Enabled         =   0   'False
      Height          =   5040
      Left            =   75
      TabIndex        =   0
      Top             =   795
      Width           =   3075
      Begin VB.PictureBox Pic_ 
         BorderStyle     =   0  'None
         Height          =   540
         Left            =   30
         ScaleHeight     =   540
         ScaleWidth      =   3000
         TabIndex        =   52
         Top             =   4455
         Width           =   3000
         Begin VB.CommandButton cmdSaveFormat 
            Caption         =   "..."
            Height          =   285
            Left            =   2310
            TabIndex        =   54
            Top             =   150
            Width           =   540
         End
         Begin VB.ComboBox cmbHW 
            Height          =   330
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   120
            Width           =   1995
         End
      End
      Begin mymoviecatalog.ShowImage scCover 
         Height          =   3225
         Left            =   390
         TabIndex        =   14
         Top             =   630
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   5689
         Picture         =   "frmMain.frx":23FF
         BorderStyle     =   0
         BackColor       =   -2147483636
      End
      Begin VB.Label lblSize 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Caption         =   "n.a"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   150
         TabIndex        =   50
         Top             =   4185
         Width           =   2730
      End
      Begin VB.Label lblHW 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "n.a"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   3885
         Width           =   2295
      End
      Begin VB.Image imgbackg 
         Height          =   3930
         Left            =   150
         Picture         =   "frmMain.frx":2CCF
         Top             =   255
         Width           =   2730
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "This is a freeware tool to find the info of your favorite movies! (c) 2009 by Salvo cortesiano."
      Height          =   420
      Left            =   4935
      TabIndex        =   3
      Top             =   165
      Width           =   6090
   End
   Begin VB.Label lblsTit 
      BackStyle       =   0  'Transparent
      Caption         =   "Movies Catalog Maker v1.0.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   165
      Width           =   4320
   End
   Begin VB.Label lblsTit 
      BackStyle       =   0  'Transparent
      Caption         =   "Movies Catalog Maker v1.0.2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Index           =   0
      Left            =   795
      TabIndex        =   1
      Top             =   195
      Width           =   4320
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmMain.frx":3EE6
      Top             =   105
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000003&
      Height          =   765
      Left            =   -30
      Top             =   -30
      Width           =   11220
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///
Option Explicit

Private statusRequest As Long
Private CancelSearch As Boolean
Private strFilesPath As String
Private tmpTitleMovie As String
Private Sub cmdCancel_Click()
    picList.Visible = False
    lstUrl.Tag = "cancel"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    Dim i As Integer
    Dim tmp As String
    For i = 0 To 14
        If txtFileds(i).Text <> "n.a" Then
            If i = 0 Then
                tmp = tmp & "Movie Title: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 1 Then
                tmp = tmp & "Direct Link Movie: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 2 Then
                tmp = tmp & "Original Title: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 3 Then
                tmp = tmp & "Country: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 4 Then
                tmp = tmp & "Year: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 5 Then
                tmp = tmp & "Gender: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 6 Then
                tmp = tmp & "Duration: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 7 Then
                tmp = tmp & "Directed By: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 8 Then
                tmp = tmp & "Official Site: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 9 Then
                tmp = tmp & "Cast: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 10 Then
                tmp = tmp & "Production: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 11 Then
                tmp = tmp & "Distribution: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 12 Then
                tmp = tmp & "Date Out: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 13 Then
                tmp = tmp & "Plot: " & txtFileds(i).Text & vbCrLf
            ElseIf i = 14 Then
                tmp = tmp & "Link of Trailer: " & txtFileds(i).Text & vbCrLf
            End If
        End If
    Next i
    Open strFilesPath + tmpTitleMovie + ".txt" For Output As #1
        Print #1, tmp
        MsgBox "The all Data of the Movie: " & txtFileds(0).Text & ", exported success!", vbInformation, App.Title
    Close #1
    If MsgBox("Copy the Data to Clipboard?", vbYesNo + vbQuestion, "Copy to Clipboard") = vbYes Then
        Clipboard.Clear
        Clipboard.SetText tmp
        MsgBox "The all Data of the Movie: " & txtFileds(0).Text & ", success to Clipboard!", vbInformation, App.Title
    End If
    tmp = Empty
End Sub

Private Sub cmdSearch_Click()
    If SearchMovie(cmbListServer.List(cmbListServer.ListIndex), txtTitle.Text) Then:
End Sub

Private Sub cmdSelect_Click()
    If lstUrl.List(lstUrl.ListIndex) = Empty Then
            MsgBox "Select one Link from the list pls!", vbExclamation, App.Title
        Exit Sub
    End If
    lstUrl.Tag = lstUrl.List(lstUrl.ListIndex)
    picList.Visible = False
End Sub

Private Sub Form_Initialize()
    '/// Init Controls XP/Vista Manifest
    '/// *****************************************************************
    If InitCommonControlsVB() Then:
    If GetDefCover() Then: _
    If FileExists(App.Path + "\cover_.jpg") Then scCover.loadimg App.Path + "\cover_.jpg"
End Sub

Private Sub Form_Load()
    cmbListServer.ListIndex = 0
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set frmMain = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub Inet_StateChanged(ByVal State As Integer)
    statusRequest = State
End Sub


Private Sub lbls_Click(Index As Integer)
    Dim OpenUrl As Integer
    On Local Error Resume Next
    If Index = 8 Then
        If Mid$(txtFileds(8).Text, 1, 7) = "http://" Then
            If MsgBox("Open the browser at: " & txtFileds(8).Text & "?", vbYesNo + _
            vbQuestion, "Open-" & txtFileds(8).Text) = vbYes Then _
            OpenUrl = ShellExecute(Me.hWnd, "Open", txtFileds(8).Text, "", App.Path, WINDOWSSTATE.OPEN_NORMAL)
        End If
    ElseIf Index = 1 Then
        If Mid$(txtFileds(1).Text, 1, 7) = "http://" Then
            If MsgBox("Open the browser at: " & txtFileds(1).Text & "?", vbYesNo + _
            vbQuestion, "Open-" & txtFileds(1).Text) = vbYes Then _
            OpenUrl = ShellExecute(Me.hWnd, "Open", txtFileds(1).Text, "", App.Path, WINDOWSSTATE.OPEN_NORMAL)
        End If
    ElseIf Index = 14 Then
         If Mid$(txtFileds(14).Text, 1, 7) = "http://" Then
            If MsgBox("Open the browser at: " & txtFileds(14).Text & "?", vbYesNo + _
            vbQuestion, "Open-" & txtFileds(14).Text) = vbYes Then _
            OpenUrl = ShellExecute(Me.hWnd, "Open", txtFileds(14).Text, "", App.Path, WINDOWSSTATE.OPEN_NORMAL)
        End If
    End If
End Sub

Private Sub lstUrl_DblClick()
    On Error Resume Next
    lstUrl.Tag = lstUrl.List(lstUrl.ListIndex)
    picList.Visible = False
End Sub

Private Sub txtTitle_Change()
    If Len(txtTitle.Text) > 0 Then
        cmdSearch.Enabled = True
    Else
        cmdSearch.Enabled = False
    End If
End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
    If Len(txtTitle.Text) > 0 Then
        cmdSearch.Enabled = True
    Else
        cmdSearch.Enabled = False
    End If
End Sub


Private Sub txtTitle_LostFocus()
    If Len(txtTitle.Text) > 0 Then
        cmdSearch.Enabled = True
        txtTitle.Text = MakeProper(txtTitle.Text)
    Else
        cmdSearch.Enabled = False
    End If
End Sub



Private Function SearchMovie(strURL As String, strTitle As String) As Boolean
    Dim strData As String
    Dim sURL As String
    Dim objLink As HTMLLinkElement
    Dim objMSHTML As New MSHTML.HTMLDocument
    Dim objDocument As MSHTML.HTMLDocument
    Dim TempArray(20) As String
    Dim FSO As New FileSystemObject
    Dim Fs As File
    Dim sT As String
    Dim tmp As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim dib As Long
    Dim i As Integer
    On Error GoTo ErrorHandler
    
    '/// Disable frame
    Frame2.Enabled = False
    frames.Enabled = False
    
    '/// Concatenate title fom (" ") to ("+")
    strTitle = Replace(strTitle, " ", "+")
    
    '/// Parsing title and convert special chars to HTML char
    strTitle = ParsingTitle(strTitle)
    If strTitle = Empty Then
                    MsgBox "Error to Parsing title!!", vbExclamation, App.Title
                Frame2.Enabled = True
            SearchMovie = False
        Exit Function
    End If
    
    '/// Call the search URL
    Select Case strURL
        Case "http://filmup.leonardo.it/"
            sURL = strURL & "cgi-bin/search.cgi?ps=10&fmt=long&q=" & strTitle & "&ul=&m=bool&wf=0020&wm=wrd&sy=0&x=33&y=10"
            strData = Inet.OpenUrl(sURL)
        Case 1
            
    End Select
    
    '/// Error request!! No Connection?!!
    If statusRequest = "11" Then
        SearchMovie = False
                MsgBox "This operation at this time is not possible!" & vbCrLf _
                & "Response code from Inet returned (11)!", vbExclamation, App.Title
            Frame2.Enabled = True
        Exit Function
    End If
    
    '/// If movie not found return FALSE
    If InStr(strData, "<small>Trovati <b>1") = 0 Then
                    MsgBox "Movie: " & txtTitle.Text & ", not found in:" & vbCrLf _
                    & strURL & "...", vbExclamation, App.Title
                Frame2.Enabled = True
            SearchMovie = False
        Exit Function
    End If
    
    '/// Empty TAG List
    lstUrl.Tag = Empty
    
    '/// Empty Firlds
    For i = 0 To 14: txtFileds(i).Text = "n.a": Next i
    
    '/// Reset labels
    lblSize.Caption = "n.a"
    lblHW.Caption = "n.a"
    
    cmdExport.Enabled = False
    
    '/// Extract All Links
    lblStatus.Caption = "Gettting document via HTTP..."
    DoEvents
    '/// This function is only available with Internet Explorer 5 > and later
    Set objDocument = objMSHTML.createDocumentFromUrl(sURL, vbNullString)
    lblStatus.Caption = "Getting and parsing HTML document..."
    DoEvents
    While objDocument.readyState <> "complete"
        DoEvents
    Wend
    lblStatus.Caption = "Document completed..."
    DoEvents
    lblStatus.Caption = "Extracting links..."
    '/// Parsing All Links
    lstUrl.Clear
    '/// Parsing potential Links
    For Each objLink In objDocument.links
        If InStr(objLink, "http://filmup.leonardo.it/sc_") Then lstUrl.AddItem objLink
        lblStatus.Caption = "Extracted " & objLink & "..."
        DoEvents
    Next
    
    '/// Exit Parsing URL
    lblStatus.Caption = "All list URL Done..."
    
    '/// Repeat until the TAG of lstURL is empty or contains "cancel" ;)
    If lstUrl.ListCount > 0 Then
        picList.Visible = True
    Do
        DoEvents
        If lstUrl.Tag <> Empty Then Exit Do
    Loop
    End If
    
    '/// Choise one Link?
    If lstUrl.Tag = "cancel" Then
        MsgBox "Search aborted by user!", vbInformation, App.Title
                Frame2.Enabled = True
            SearchMovie = False
        Exit Function
    End If
    
    '/// Display the Link of movie
    txtFileds(1).Text = lstUrl.Tag
    
    '/// Go to link and Get ALL Info of the Movie
    strData = Inet.OpenUrl(lstUrl.Tag)
    
    '/// Error request!! No Connection?!!
    If statusRequest = "11" Then
        SearchMovie = False
                MsgBox "This operation at this time is not possible!" & vbCrLf _
                & "Response code from Inet() returned (11)!" & vbCrLf _
                & "Request URL: " & lstUrl.Tag, vbExclamation, App.Title
            Frame2.Enabled = True
        Exit Function
    End If
    
    '/// Clear Cover
    If GetDefCover() Then: _
        If FileExists(App.Path + "\cover_.jpg") Then scCover.loadimg App.Path + "\cover_.jpg"
    
    '/// Clear URL
    lstUrl.Tag = Empty
    
    '/// The link of Triler exist?
    Set objDocument = objMSHTML.createDocumentFromUrl(txtFileds(1).Text, vbNullString)
    lblStatus.Caption = "Getting and parsing HTML document URL-Trailer..."
    DoEvents
    While objDocument.readyState <> "complete"
        DoEvents
    Wend
    lblStatus.Caption = "Document completed..."
    DoEvents
    lstUrl.Clear
    lblStatus.Caption = "Extracting links..."
    '/// Parsing potential Links if contains Trailers
    For Each objLink In objDocument.links
        If InStr(objLink, "http://filmup.leonardo.it/trailers/") Then lstUrl.AddItem objLink
        lblStatus.Caption = "Extracted " & objLink & "-trailers"
        DoEvents
    Next
    lblStatus.Caption = "Document completed..."
    
    '/// Repeat until the TAG of lstURL is empty or contains "cancel" ;)
    If lstUrl.ListCount > 0 Then
        picList.Visible = True
    Do
        DoEvents
        If lstUrl.Tag <> Empty Then Exit Do
    Loop
    End If
    
    '/// Choise one Link?
    If lstUrl.Tag = "cancel" Then
        TempArray(20) = "n.a"
    Else
        TempArray(20) = lstUrl.Tag
    End If
    
    '/// Clear TAG
    lstUrl.Tag = Empty
    
    '/// Extract the exat Title
    txtFileds(0).Text = MakeProper(ParsingString(1, strData, "<title>FilmUP - Scheda: ", "</title>", vbTextCompare))
    
    '/// Create the folder Movie and the Sub-Folder of the Movie title if not exist
    tmpTitleMovie = txtFileds(0).Text
    
    '/// Remove the special chars in the Title of movie
    tmpTitleMovie = Replace(tmpTitleMovie, ":", "-")
    tmpTitleMovie = Replace(tmpTitleMovie, "\", "-")
    tmpTitleMovie = Replace(tmpTitleMovie, "/", "-")
    
    '/// Create sub-Folder
    If Not FSO.FolderExists(App.Path + "\movies\" + tmpTitleMovie) Then
        ' .... Create folder .torrent
        If MakeDirectory(App.Path + "\movies\" + tmpTitleMovie) = False Then:
        ' .... Until display the Error now, because if the Folder exist return a Error ;)
    End If
    '/// Last verify
    If FSO.FolderExists(App.Path + "\movies\" + tmpTitleMovie) Then
        strFilesPath = App.Path + "\movies\" + tmpTitleMovie & "\"
    Else
        strFilesPath = App.Path + "\"
    End If
    
    
    '///|\\\---------------------------------------------------------ORIGINAL TITLE
    
    
    '/// Extract Original Title
    If InStr(strData, "Titolo originale:&nbsp;</font></td>") Then
    pos1 = InStr(strData, "Titolo originale:&nbsp;</font></td>")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(1) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(1) = "n.a"
    End If
    End If
    
    txtFileds(2).Text = MakeProper(TempArray(1))
    
    
    '///|\\\---------------------------------------------------------COUNTRY
    
    
    '/// Extract Country
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Nazione:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Nazione:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(2) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(2) = "n.a"
    End If
    End If
    
    txtFileds(3).Text = MakeProper(TempArray(2))
    
    
    '///|\\\---------------------------------------------------------YEAR
    
    
    '/// Extract Year
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Anno:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Anno:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(3) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(3) = "n.a"
    End If
    End If
    
    txtFileds(4).Text = MakeProper(TempArray(3))
    
    
    '///|\\\---------------------------------------------------------GENDER
    
    
    '/// Extract Gender
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Genere:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Genere:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(4) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(4) = "n.a"
    End If
    End If
    
    txtFileds(5).Text = MakeProper(TempArray(4))
    
    
    '///|\\\---------------------------------------------------------DURATION
    
    
    '/// Extract Duration
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Durata:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Durata:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(5) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(5) = "n.a"
    End If
    End If
    
    txtFileds(6).Text = TempArray(5)
    
    
    '///|\\\---------------------------------------------------------DIRECTED BY
    
    
    '/// Extract Director
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Regia:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Regia:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(6) = Mid$(strData, pos1, pos2 - pos1)
    
    If InStr(TempArray(6), "href") > 0 Then
        pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Regia:")
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" _
        & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</td>", vbTextCompare)
        TempArray(6) = Mid$(strData, pos1, pos2 - pos1)
        TempArray(6) = Replace(TempArray(6), "</font>", "")
        TempArray(6) = Replace(TempArray(6), "</a>, ", ", ")
        tmp = TempArray(6)
        For i = 0 To 3
            tmp = SimpleHTMLRep(TempArray(6), "<", ">", "")
            If tmp <> "n.a" Then TempArray(6) = tmp
        Next i
        
    End If
    
    Else
        TempArray(6) = "n.a"
    End If
    End If
    
    txtFileds(7).Text = TempArray(6)
    
    
    '///|\\\---------------------------------------------------------OFFICIAL SITE

    
    '/// Official Site
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Sito ufficiale:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Sito ufficiale:")
    If pos1 > 0 Then
        
    pos1 = InStr(strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">&nbsp")
    If pos1 > 0 Then
        TempArray(7) = "n.a"
    Else
        pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Sito ufficiale:") '// da togliere?
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & "><a class=""" & "filmup""" & " href=", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & "><a class=""" & "filmup""" & " href=") + 1
        pos2 = InStr(pos1 + 1, strData, "target=""" & "link""" & ">", vbTextCompare) - 2
        TempArray(7) = Mid$(strData, pos1, pos2 - pos1)
    End If
    
    If Len(TempArray(7)) = 0 Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Sito ufficiale:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & "><a class=""" & "filmup""" & " target=""" & "link""" & " href=", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & "><a class=""" & "filmup""" & " target=""" & "link""" & " href=") + 1
        pos2 = InStr(pos1 + 1, strData, ">", vbTextCompare) - 1
        TempArray(7) = Mid$(strData, pos1, pos2 - pos1)
    End If
    End If
    Else
        TempArray(7) = "n.a"
    End If
    End If
    
    txtFileds(8).Text = TempArray(7)
    
    
    '///|\\\---------------------------------------------------------THE CAST
    
    
    '/// Extract the Cast
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Cast:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Cast:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" _
        & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</td>", vbTextCompare)
        TempArray(8) = Mid$(strData, pos1, pos2 - pos1)
        TempArray(8) = Replace(TempArray(8), "</font>", "")
        TempArray(8) = Replace(TempArray(8), "</a>, ", ", ")
        tmp = TempArray(8)
        For i = 0 To 5
            tmp = SimpleHTMLRep(TempArray(8), "<", ">", "")
            If tmp <> "n.a" Then TempArray(8) = tmp
        Next i
    Else
        TempArray(8) = "n.a"
    End If
    End If

    txtFileds(9).Text = TempArray(8)
    
    
    '///|\\\---------------------------------------------------------PRODUCTION
    
    
    '/// Extract the Production
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Produzione:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Produzione:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(10) = Mid$(strData, pos1, pos2 - pos1)
        
        If InStr(TempArray(10), "href") > 0 Then
        pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Produzione:")
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" _
        & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</td>", vbTextCompare)
        TempArray(10) = Mid$(strData, pos1, pos2 - pos1)
        TempArray(10) = Replace(TempArray(10), "</font>", "")
        TempArray(10) = Replace(TempArray(10), "</a>, ", ", ")
        tmp = TempArray(10)
        For i = 0 To 3
            tmp = SimpleHTMLRep(TempArray(10), "<", ">", "")
            If tmp <> "n.a" Then TempArray(10) = tmp
        Next i
        
    End If
    
    Else
        TempArray(10) = "n.a"
    End If
    End If
    
    txtFileds(10).Text = TempArray(10)
    
    
    '///|\\\---------------------------------------------------------DISTRIBUTION
    
    
    '/// Extract Distribution
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Distribuzione:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Distribuzione:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "target=""" & "_blank""" & ">", vbTextCompare) _
        + Len("target=""" & "_blank""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</a></font></td>", vbTextCompare)
        TempArray(11) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(11) = "n.a"
    End If
    End If
    
    If TempArray(11) = "&nbsp;" Then TempArray(11) = "n.a"
    txtFileds(11).Text = TempArray(11)
    
    
    '///|\\\---------------------------------------------------------DATE MOVIE OUT
    
    
    '/// Extract date Out
    If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Data di uscita:") Then
    pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Data di uscita:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(12) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(12) = "n.a"
    End If
    Else
        If InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Uscita prevista:") Then
        pos1 = InStr(strData, "<td valign=""" & "top""" & " nowrap><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Uscita prevista:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">", vbTextCompare) _
        + Len("<td valign=""" & "top""" & "><font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">")
        pos2 = InStr(pos1 + 1, strData, "</font></td>", vbTextCompare)
        TempArray(12) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(12) = "n.a"
    End If
    End If
    
    End If
    
    If TempArray(12) = "&nbsp;" Then TempArray(12) = "n.a"
    TempArray(12) = Replace(TempArray(12), "<br />", " ")
    txtFileds(12).Text = TempArray(12)
    
    
    '///|\\\---------------------------------------------------------THE PLOTS
    
    
    '/// Extract Plot
    If InStr(strData, "<font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Trama:") Then
    pos1 = InStr(strData, "<font face=""" & "arial, helvetica""" & " size=""" & "2""" & ">Trama:")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "<br>", vbTextCompare) + 4
        pos2 = InStr(pos1 + 1, strData, "</font><br>", vbTextCompare)
        TempArray(13) = Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(13) = "n.a"
    End If
    End If
    
    txtFileds(13).Text = TempArray(13)
    
    '/// Link of the Triler
    txtFileds(14) = TempArray(20)
    
    '///|\\\---------------------------------------------------------LINK OF THE COVER
    
    
    '/// Extract Link Cover
    If InStr(strData, "<td align=""" & "left""" & "><a class=""" & "filmup""" & " href=""" & "posters/locp/") Then
    pos1 = InStr(strData, "<td align=""" & "left""" & "><a class=""" & "filmup""" & " href=")
    If pos1 > 0 Then
        pos1 = InStr(pos1 + 1, strData, "href=", vbTextCompare) + 6
        pos2 = InStr(pos1 + 1, strData, " target=", vbTextCompare) - 1
        TempArray(14) = "http://filmup.leonardo.it/" & Mid$(strData, pos1, pos2 - pos1)
    Else
        TempArray(14) = "n.a"
    End If
    End If
    
    If TempArray(14) <> "n.a" Then
    '/// Go page of Poster
    strData = Inet.OpenUrl(TempArray(14))
    If statusRequest = "11" Then
            Frame2.Enabled = True
            MsgBox "This operation at this time is not possible!" & vbCrLf _
                & "Response code from Inet returned (11)!", vbExclamation, App.Title
        Exit Function
    End If
    End If
    
    '/// Download Poster
    TempArray(15) = "http://filmup.leonardo.it/posters/loc" & ParsingString(1, strData, _
    "<img src=""" & "../loc", " width=", vbTextCompare)
    
    '/// Remuve special TAG
    TempArray(15) = Replace(TempArray(15), """", "")
    TempArray(15) = Replace(TempArray(15), " ", "")
    TempArray(15) = Replace(TempArray(15), "=", "")
    TempArray(15) = Replace(TempArray(15), "alt", "")
    TempArray(15) = Replace(TempArray(15), "/>", "")
    TempArray(15) = Replace(TempArray(15), "\<", "")
    
    '/// Remuve special CHARS from Title
    tmpTitleMovie = txtFileds(0).Text
    tmpTitleMovie = Replace(tmpTitleMovie, ":", "-")
    tmpTitleMovie = Replace(tmpTitleMovie, "\", "-")
    tmpTitleMovie = Replace(tmpTitleMovie, "/", "-")
    
    If FileExists(strFilesPath + tmpTitleMovie + ".jpg") Then
        '/// Display the Cover
        scCover.loadimg strFilesPath + tmpTitleMovie + ".jpg"
        '/// Get the Size
        Set Fs = FSO.GetFile(strFilesPath + tmpTitleMovie + ".jpg")
        lblSize.Caption = GetSizeBytes(Fs.Size, DISP_BYTES_SHORT)
        '/// Get the Height and Width
        If FileExists(App.Path + "\FreeImage.dll") Then
            dib = FreeImage_LoadEx(strFilesPath + tmpTitleMovie + ".jpg")
            lblHW.Caption = FreeImage_GetWidth(dib) & "x" & FreeImage_GetHeight(dib)
            '/// Adding other H and W
            cmbHW.Clear
            cmbHW.AddItem FreeImage_GetWidth(dib) & "x" & FreeImage_GetHeight(dib)
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 & "x" & FreeImage_GetHeight(dib) / 2
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 / 2 & "x" & FreeImage_GetHeight(dib) / 2 / 2
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 / 2 / 2 & "x" & FreeImage_GetHeight(dib) / 2 / 2 / 2
            cmbHW.ListIndex = 0
            Call FreeImage_Unload(dib)
        End If
    Else
    
    '/// If download Cover = SUCCESS
    If DownloadFile(TempArray(15), strFilesPath + tmpTitleMovie + ".jpg", , , tmpTitleMovie + ".jpg") = True Then
        If FileExists(strFilesPath + tmpTitleMovie + ".jpg") Then
            frames.Enabled = True
            '/// Display the Cover
            scCover.loadimg strFilesPath + tmpTitleMovie + ".jpg"
            '/// Get the Size
            Set Fs = FSO.GetFile(strFilesPath + tmpTitleMovie + ".jpg")
            lblSize.Caption = GetSizeBytes(Fs.Size, DISP_BYTES_SHORT)
        Else
            If GetDefCover() Then: _
            If FileExists(App.Path + "\cover_.jpg") Then scCover.loadimg App.Path + "\cover_.jpg"
            frames.Enabled = False
        End If
        '/// Get the Height and Width
        If FileExists(App.Path + "\FreeImage.dll") Then
            dib = FreeImage_LoadEx(strFilesPath + tmpTitleMovie + ".jpg")
            lblHW.Caption = FreeImage_GetWidth(dib) & "x" & FreeImage_GetHeight(dib)
            '/// Adding other H and W
            cmbHW.Clear
            cmbHW.AddItem FreeImage_GetWidth(dib) & "x" & FreeImage_GetHeight(dib)
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 & "x" & FreeImage_GetHeight(dib) / 2
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 / 2 & "x" & FreeImage_GetHeight(dib) / 2 / 2
            cmbHW.AddItem FreeImage_GetWidth(dib) / 2 / 2 / 2 & "x" & FreeImage_GetHeight(dib) / 2 / 2 / 2
            cmbHW.ListIndex = 0
            Call FreeImage_Unload(dib)
        End If
    Else
        If GetDefCover() Then: _
        If FileExists(App.Path + "\cover_.jpg") Then scCover.loadimg App.Path + "\cover_.jpg"
        frames.Enabled = False
    End If
    
    End If
    
    '///|\\\---------------------------------------------------------LINK OF THE TRAILER
    
    If FileExists(strFilesPath + tmpTitleMovie + ".mov") Then
    
    Else
    '/// Go to link of Movie Trailer if exist
    If ChkDownloadMovie.Value = 1 Then
    If TempArray(20) <> "n.a" Then
        DoEvents
        lblStatus.Caption = "Download movie Trailer..."
        TempArray(20) = Replace(TempArray(20), ".shtml", ".mov")
        TempArray(20) = Replace(TempArray(20), "filmup", "mediafilmup")
        DoEvents
        If DownloadFile(TempArray(20), strFilesPath + tmpTitleMovie + ".mov", , , tmpTitleMovie + ".mov") = True Then
            If FileExists(strFilesPath + tmpTitleMovie + ".mov") Then
                DoEvents
                    lblStatus.Caption = "Download movie Trailer...Completate Ok!"
                End If
            End If
        End If
    End If
    End If
    
    '///|\\\--------------------------------------------------------END CODE! _._._._ ///
    '/// Release FileSystemObject
    Set FSO = Nothing
    Set Fs = Nothing
    
    '/// Enabled All
    Frame2.Enabled = True
    
    '/// Empty all Variables
    i = 0
    For i = 0 To 20: TempArray(i) = Empty: Next
    
    strData = Empty
    
    lblStatus.Caption = "Info movie loaded success Ok!"
    
    cmdExport.Enabled = True
    
    '/// Return Success...
    SearchMovie = True
Exit Function

ResetAllFrames:
    lblStatus.Caption = "Reset after Error OK!"
    strData = Empty
    i = 0
    For i = 0 To 15: TempArray(i) = Empty: Next
    i = 0
    For i = 0 To 14: txtFileds(i).Text = "n.a": Next i
    lblSize.Caption = "n.a"
    lblHW.Caption = "n.a"
    If GetDefCover() Then: _
    If FileExists(App.Path + "\cover_.jpg") Then scCover.loadimg App.Path + "\cover_.jpg"
    Frame2.Enabled = True
Exit Function

ErrorHandler:
    If ChkError.Value = 1 Then
        Resume Next
    Else
    MsgBox "Error #" & Err.Number & "." & Err.Description, vbExclamation, App.Title
        SearchMovie = False
        GoTo ResetAllFrames
    End If
    Err.Clear
End Function
Private Function ParsingTitle(strString As String) As String
    On Local Error GoTo ParsingError
        '/// SPECIAL SMAL CHARS
        '-----------------------------------------
        strString = Replace(strString, "ê", "%EA")
        strString = Replace(strString, "ë", "%EB")
        strString = Replace(strString, "ì", "%EC")
        strString = Replace(strString, "í", "%ED")
        strString = Replace(strString, "î", "%EE")
        strString = Replace(strString, "ï", "%FF")
        strString = Replace(strString, "à", "%E0")
        strString = Replace(strString, "á", "%E1")
        strString = Replace(strString, "â", "%E2")
        strString = Replace(strString, "ã", "%E3")
        strString = Replace(strString, "ä", "%E4")
        strString = Replace(strString, "å", "%E5")
        strString = Replace(strString, "æ", "%E6")
        strString = Replace(strString, "ç", "%E7")
        strString = Replace(strString, "è", "%E8")
        strString = Replace(strString, "é", "%E9")
        strString = Replace(strString, "ð", "%F0")
        strString = Replace(strString, "ñ", "%F1")
        strString = Replace(strString, "ò", "%F2")
        strString = Replace(strString, "ó", "%F3")
        strString = Replace(strString, "ô", "%F4")
        strString = Replace(strString, "õ", "%F5")
        strString = Replace(strString, "ö", "%F6")
        strString = Replace(strString, "÷", "%F7")
        strString = Replace(strString, "ø", "%F8")
        strString = Replace(strString, "ù", "%F9")
        strString = Replace(strString, "ú", "%FA")
        strString = Replace(strString, "û", "%FB")
        strString = Replace(strString, "ü", "%FC")
        strString = Replace(strString, "ý", "%FD")
        strString = Replace(strString, "þ", "%FE")
        strString = Replace(strString, "ÿ", "%FF")
        '/// SPECIAL SIMBOLS
        '-----------------------------------------
        'strString = Replace(strString, "+", "%2B")
        strString = Replace(strString, ",", "%2C")
        strString = Replace(strString, "-", "%2D")
        strString = Replace(strString, ".", "%2E")
        strString = Replace(strString, "/", "%2F")
        strString = Replace(strString, ":", "%3A")
        strString = Replace(strString, ";", "%3B")
        strString = Replace(strString, "<", "%3C")
        strString = Replace(strString, "=", "%3D")
        strString = Replace(strString, ">", "%3E")
        strString = Replace(strString, "?", "%3F")
        strString = Replace(strString, "[", "%5B")
        strString = Replace(strString, "\", "%5C")
        strString = Replace(strString, "]", "%5D")
        strString = Replace(strString, "^", "%5E")
        strString = Replace(strString, "_", "%5F")
        strString = Replace(strString, "`", "%60")
        strString = Replace(strString, "{", "%7B")
        strString = Replace(strString, "|", "%7C")
        strString = Replace(strString, "}", "%7D")
        strString = Replace(strString, "~", "%7E")
        strString = Replace(strString, "", "%7F")
        strString = Replace(strString, "Š", "%8A")
        strString = Replace(strString, "‹", "%8B")
        strString = Replace(strString, "Œ", "%8C")
        strString = Replace(strString, "ž", "%8E")
        strString = Replace(strString, "š", "%9A")
        strString = Replace(strString, "›", "%9B")
        strString = Replace(strString, "œ", "%9C")
        strString = Replace(strString, "Ÿ", "%9F")
        strString = Replace(strString, "¡", "%A1")
        strString = Replace(strString, "¢", "%A2")
        strString = Replace(strString, "£", "%A3")
        strString = Replace(strString, "¤", "%A4")
        strString = Replace(strString, "¥", "%A5")
        strString = Replace(strString, "¦", "%A6")
        strString = Replace(strString, "§", "%A7")
        strString = Replace(strString, "¨", "%A8")
        strString = Replace(strString, "©", "%A9")
        strString = Replace(strString, "ª", "%AA")
        strString = Replace(strString, "«", "%AB")
        strString = Replace(strString, "¬", "%AC")
        strString = Replace(strString, "®", "%AE")
        strString = Replace(strString, "¯", "%AF")
        'strString = Replace(strString, "°", "%B0")
        strString = Replace(strString, "±", "%B1")
        strString = Replace(strString, "²", "%B2")
        strString = Replace(strString, "³", "%B3")
        strString = Replace(strString, "´", "%B4")
        strString = Replace(strString, "µ", "%B5")
        strString = Replace(strString, "¶", "%B6")
        strString = Replace(strString, "·", "%B7")
        strString = Replace(strString, "¸", "%B8")
        strString = Replace(strString, "¹", "%B9")
        strString = Replace(strString, "º", "%BA")
        strString = Replace(strString, "»", "%BB")
        strString = Replace(strString, "¼", "%BC")
        strString = Replace(strString, "½", "%BD")
        strString = Replace(strString, "¾", "%BE")
        strString = Replace(strString, "¿", "%BF")
        '/// SPECIAL BIG CHARS
        '-----------------------------------------
        strString = Replace(strString, "À", "%C0")
        strString = Replace(strString, "Á", "%C1")
        strString = Replace(strString, "Â", "%C2")
        strString = Replace(strString, "Ã", "%C3")
        strString = Replace(strString, "Ä", "%C4")
        strString = Replace(strString, "Å", "%C5")
        strString = Replace(strString, "Æ", "%C6")
        strString = Replace(strString, "Ç", "%C7")
        strString = Replace(strString, "È", "%C8")
        strString = Replace(strString, "É", "%C9")
        strString = Replace(strString, "Ê", "%CA")
        strString = Replace(strString, "Ë", "%CB")
        strString = Replace(strString, "Ì", "%CC")
        strString = Replace(strString, "Í", "%CD")
        strString = Replace(strString, "Î", "%CE")
        strString = Replace(strString, "Ï", "%CF")
        strString = Replace(strString, "Ð", "%D0")
        strString = Replace(strString, "Ñ", "%D1")
        strString = Replace(strString, "Ò", "%D2")
        strString = Replace(strString, "Ó", "%D3")
        strString = Replace(strString, "Ô", "%D4")
        strString = Replace(strString, "Õ", "%D5")
        strString = Replace(strString, "Ö", "%D6")
        strString = Replace(strString, "×", "%D7")
        strString = Replace(strString, "Ø", "%D8")
        strString = Replace(strString, "Ù", "%D9")
        strString = Replace(strString, "Ú", "%DA")
        strString = Replace(strString, "Û", "%DB")
        strString = Replace(strString, "Ü", "%DC")
        strString = Replace(strString, "Ý", "%DD")
        strString = Replace(strString, "Þ", "%DE")
        strString = Replace(strString, "ß", "%DF")
        ParsingTitle = strString
    Exit Function
ParsingError:
    ParsingTitle = Empty
End Function

Private Function ParsingString(ByVal Start As Long, Data As String, StartString As String, EndString As String, _
    Optional ByVal CompareMethod As VbCompareMethod = vbBinaryCompare) As String
    Dim lonStart As Long, lonEnd As Long
    On Local Error Resume Next
    lonStart = InStr(Start, Data, StartString, CompareMethod)
    If lonStart > 0 Then
        lonStart = lonStart + Len(StartString)
        lonEnd = InStr(lonStart, Data, EndString, CompareMethod)
        If lonEnd > 0 Then
            ParsingString = Mid$(Data, lonStart, lonEnd - lonStart)
        End If
    End If
End Function

Private Function DownloadFile(strURL As String, strDestination As String, Optional UserName As String = Empty, Optional password As String = Empty, Optional strFileName As String = Empty) As Boolean

Const CHUNK_SIZE As Long = 1024
Const ROLLBACK As Long = 4096

Dim bData() As Byte
Dim blnResume As Boolean
Dim intFile As Integer
Dim lngBytesReceived As Long
Dim lngFileLength As Long
Dim lngX
Dim sglLastTime As Single
Dim sglRate As Single
Dim sglTime As Single
Dim strFile As String
Dim strHeader As String
Dim strHost As String

On Local Error GoTo InternetErrorHandler

strFile = ReturnFileOrFolder(strDestination, True)
strHost = ReturnFileOrFolder(strURL, True, True)

StartDownload:
If blnResume Then
    lblStatus.Caption = "Resuming download..."
    lngBytesReceived = lngBytesReceived - ROLLBACK
    If lngBytesReceived < 0 Then lngBytesReceived = 0
Else
    lblStatus.Caption = "Retrive Poster..."
End If
DoEvents

With Inet
    .url = strURL
    .UserName = UserName
    .password = password
    .Execute , "GET", , "Range: bytes=" & CStr(lngBytesReceived) & "-" & vbCrLf
    While .StillExecuting
        DoEvents
    Wend
    
    lblStatus.Caption = "Download Poster..."

    strHeader = .GetHeader
    
End With

Select Case Mid$(strHeader, 10, 3)
    Case "200"
        If blnResume Then
            Kill strDestination
            If MsgBox("Impossibile riesumare il Download." & vbCr & vbCr & "Vuoi comunque continuare?", _
                     vbExclamation + vbYesNo, "Resume Download") = vbYes Then
                    blnResume = False
                Else
                    CancelSearch = True
                    GoTo ExitDownload
                End If
            End If
    Case "206"  ' 206=Contenuto Parziale
    Case "204"
        MsgBox "Niente da scaricare!", vbInformation, "Nessun Download"
        CancelSearch = True
        GoTo ExitDownload
    Case "401"
        MsgBox "Autorizzazione (negata) Download del file fallito!", vbCritical, "Non autorizzato"
        CancelSearch = True
        GoTo ExitDownload
    Case "404"  ' File non trovato
        MsgBox "File " & """" & strFileName & """" & " non presente sul server o agiornamento non disponibile!", vbCritical, "File non trovato"
        CancelSearch = True
        GoTo ExitDownload
    Case vbCrLf
    MsgBox "Impossibile stabilire una connessione." & vbCr & vbCr & "Verificare le Impostazioni di connessione alla rete e riprovare", _
               vbExclamation, "Impossibile Connettersi"
        CancelSearch = True
        GoTo ExitDownload
    Case Else
        strHeader = Left(strHeader, InStr(strHeader, vbCr))
        If strHeader = Empty Then strHeader = "<nothing>"
        MsgBox "Il server ha risposto in questo modo:" & vbCr & vbCr & strHeader, vbCritical, "Errore Downloading File"
        CancelSearch = True
        GoTo ExitDownload
End Select
If blnResume = False Then
    sglLastTime = Timer - 1
    strHeader = Inet.GetHeader("Content-Length")
    lngFileLength = Val(strHeader)
    If lngFileLength = 0 Then
        GoTo ExitDownload
    End If
End If

DoEvents
If blnResume = False Then lngBytesReceived = 0
On Local Error GoTo FileErrorHandler
strHeader = ReturnFileOrFolder(strDestination, False)
If Dir(strHeader, vbDirectory) = Empty Then
    MkDir strHeader
End If
    intFile = FreeFile()
    Open strDestination For Binary Access Write As #3
    If blnResume Then Seek #3, lngBytesReceived + 1
    Do
    bData = Inet.GetChunk(CHUNK_SIZE, icByteArray)
    Put #3, , bData
    If CancelSearch Then Exit Do
    lngBytesReceived = lngBytesReceived + UBound(bData, 1) + 1
    sglRate = lngBytesReceived / (Timer - sglLastTime)
    sglTime = (lngFileLength - lngBytesReceived) / sglRate
    lblStatus.Caption = "Download: " & FormatTime(sglTime) & " (" & FormatFileSize(lngBytesReceived) _
    & " of " & FormatFileSize(lngFileLength) & " /" & FormatFileSize(sglRate, "###.0") & "/sec"
    DoEvents
Loop While UBound(bData, 1) > 0
Close #3
    If CancelSearch <> True Then
    End If
ExitDownload:
If lngBytesReceived = lngFileLength And CancelSearch = False Then
    lblStatus.Caption = "Download Poster Ok!"
    DownloadFile = True
    GoTo Cleanup
Else
    If CancelSearch = True Then
        lblStatus.Caption = "Error to Download Poster!"
    End If
    
    If Dir(strDestination) = Empty Then
        CancelSearch = True
    Else
        If CancelSearch = False Then
                    blnResume = True
                    GoTo StartDownload
            End If
        End If
    End If
    
    DownloadFile = False
'End If

Cleanup:
Inet.Cancel
Exit Function
InternetErrorHandler:
    If Err.Number = 9 Then Resume Next
    MsgBox "Error #" & Err.Description, vbCritical, "Errore Downloading File"
    Err.Clear
    GoTo ExitDownload
FileErrorHandler:
    MsgBox "Errore #" & Err.Number & ": " & Err.Description, vbCritical, "Error Downloading File"
    CancelSearch = True
    Err.Clear
    GoTo ExitDownload
End Function

Private Function ReturnFileOrFolder(FullPath As String, ReturnFile As Boolean, Optional IsURL As Boolean = False) As String
Dim intDelimiterIndex As Integer
On Error Resume Next
intDelimiterIndex = InStrRev(FullPath, IIf(IsURL, "/", "\"))
If intDelimiterIndex = 0 Then
    ReturnFileOrFolder = FullPath
Else
    ReturnFileOrFolder = IIf(ReturnFile, Right(FullPath, Len(FullPath) - intDelimiterIndex), Left(FullPath, intDelimiterIndex))
End If
End Function

Private Function SimpleHTMLRep(SearchString As String, FirstTag As String, SecondTag As String, Optional MyReplacement As String = "") As String
    Dim MyPos1 As Integer: Dim MyPos2 As Integer: Dim i As Integer
    Dim FirstPart As String: Dim LastPart As String
    On Local Error GoTo ErrorParsing
        MyPos1 = InStr(1, UCase(SearchString), UCase(FirstTag), 1)
        MyPos2 = InStr(MyPos1, UCase(SearchString), UCase(SecondTag), 1)
        FirstPart = Mid$(SearchString, 1, MyPos1 - 1)
        LastPart = Mid$(SearchString, MyPos2 + Len(SecondTag), Len(SearchString) - MyPos2)
    If MyReplacement > "" Then SimpleHTMLRep = FirstPart & MyReplacement & LastPart Else _
                    SimpleHTMLRep = FirstPart & LastPart
Exit Function
ErrorParsing:
        SimpleHTMLRep = "n.a"
    Err.Clear
End Function

Private Function GetDefCover() As Boolean
    Dim f As Integer: Dim b() As Byte
    On Local Error GoTo ErrorHandler
    f = FreeFile
        b = LoadResData(101, "PICTURE")
        Open App.Path + "\cover_.jpg" For Binary Access Write Shared As #f
            Put #f, , b
        Close #f
    If FileExists(App.Path + "\cover_.jpg") Then GetDefCover = True Else GetDefCover = False
Exit Function
ErrorHandler:
        GetDefCover = False
    Err.Clear
End Function
