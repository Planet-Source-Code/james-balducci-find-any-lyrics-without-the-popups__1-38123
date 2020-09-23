VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lyrics Search"
   ClientHeight    =   6315
   ClientLeft      =   150
   ClientTop       =   765
   ClientWidth     =   11445
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtList 
      Height          =   270
      Left            =   4320
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   5910
      Visible         =   0   'False
      Width           =   1020
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   6660
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6510
      Left            =   5550
      TabIndex        =   14
      Top             =   75
      Width           =   5880
      ExtentX         =   10372
      ExtentY         =   11483
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   30
      TabIndex        =   10
      Top             =   1650
      Width           =   5445
      Begin MSComDlg.CommonDialog cd 
         Left            =   4155
         Top             =   2790
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Save Lyrics to File"
      End
      Begin VB.TextBox tOldT1 
         Height          =   480
         Left            =   2580
         TabIndex        =   18
         Top             =   2505
         Visible         =   0   'False
         Width           =   1005
      End
      Begin RichTextLib.RichTextBox txtLyrics 
         Height          =   4635
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   8176
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         MousePointer    =   1
         DisableNoScroll =   -1  'True
         Appearance      =   0
         TextRTF         =   $"Form1.frx":74F2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1770
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   5445
      Begin VB.CommandButton Command3 
         Caption         =   "Show Songs"
         Default         =   -1  'True
         Height          =   420
         Left            =   105
         TabIndex        =   15
         Top             =   1080
         Width           =   1590
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   405
         Left            =   3780
         Picture         =   "Form1.frx":7569
         TabIndex        =   9
         Top             =   1110
         Width           =   1545
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   720
         TabIndex        =   6
         Top             =   645
         Width           =   4635
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   720
         TabIndex        =   5
         Top             =   240
         Width           =   4620
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   2340
         Picture         =   "Form1.frx":EA5B
         Top             =   975
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Song:"
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Artist:"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Parse"
      Height          =   315
      Left            =   5520
      TabIndex        =   3
      Top             =   4455
      Visible         =   0   'False
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox txtLoad 
      Height          =   2640
      Left            =   5940
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   3570
      Width           =   3570
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   870
      TabIndex        =   1
      Text            =   "Text4"
      Top             =   2760
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.TextBox Text3 
      Height          =   360
      Left            =   795
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   2310
      Visible         =   0   'False
      Width           =   3780
   End
   Begin VB.Frame Frame3 
      Height          =   405
      Left            =   30
      TabIndex        =   11
      Top             =   6180
      Visible         =   0   'False
      Width           =   5445
      Begin VB.Label Label3 
         Caption         =   "Ready..."
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   4485
      End
   End
   Begin VB.Label Quot 
      Caption         =   """"
      Height          =   450
      Left            =   4080
      TabIndex        =   16
      Top             =   5055
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu save 
         Caption         =   "&Save to file"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3 = "http://www.azlyrics.com/" & Left(Text1, 1) & "/" & Replace(Text1, " ", "") & ".html"
Text4 = "http://www.azlyrics.com/lyrics/" & Replace(Text1, " ", "") & "/" & Replace(Text2, " ", "") & ".html"
Label3 = "Searching..."
    On Error Resume Next
    
    Dim txt As String
    Dim b() As Byte
    
    Command1.Enabled = False
    
    b() = Inet1.OpenURL(Text4.Text, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    txtLoad.Text = txt
    Command1.Enabled = True
    Label3 = "Page Loaded"
    Command2_Click
    
    Exit Sub

End Sub

Private Sub Command2_Click()
Dim theStart
Dim theLength
On Error GoTo notFound
theStart = InStr(1, txtLoad.Text, "<B>" + UCase(Text1.Text) + " LYRICS", vbTextCompare)
'txtLoad = Right(txtLoad, theStart)
theLength = Len(txtLoad) - theStart
txtLoad = Mid(txtLoad, theStart, theLength)
'theStart = InStr(1, txtLoad.Text, Text2, vbTextCompare) + Len(Text2) + 1
'theLength = Len(txtLoad) - theStart
'txtLoad = Mid(txtLoad, theStart, theLength)
txtLoad = Replace(txtLoad.Text, "" + UCase(Text1.Text) + " LYRICS", "")
txtLoad = Replace(txtLoad.Text, "<BR><BR>", vbCrLf & "\par " & vbCrLf & "\par ")
txtLoad = Replace(txtLoad.Text, "</B>", "\b0")
txtLoad = Replace(txtLoad.Text, "<br>", vbCrLf & "\par ")
txtLoad = Replace(txtLoad.Text, "<BR>", vbCrLf & "\par ")
txtLoad = Replace(txtLoad.Text, "<FONT size=2>", "")
txtLoad = Replace(txtLoad.Text, "<B>", "\b1")
txtLoad = Replace(txtLoad.Text, " <BR>", vbCrLf & "\par ")
txtLoad = Replace(txtLoad.Text, " <br>", vbCrLf & "\par ")
txtLoad = Replace(txtLoad.Text, "<b>", "\b1")
txtLoad = Replace(txtLoad.Text, "</b>", "\b0")
txtLoad = Replace(txtLoad.Text, "<i>", "\I1")
txtLoad = Replace(txtLoad.Text, "</i>", "\I0")
txtLoad = Replace(txtLoad.Text, "<I>", "\I1")
txtLoad = Replace(txtLoad.Text, "</I>", "\I0")
theStart = InStr(1, txtLoad.Text, "[ <a href=")
txtLoad = Left(txtLoad, theStart - 1)
theStart = InStr(1, txtLoad.Text, Quot)
txtLoad = Mid(txtLoad, theStart, Len(txtLoad) - theStart)
Label3 = "Done..."
txtLyrics.TextRTF = "{" & vbCrLf & "\viewkind4\uc1\pard\f0\fs20\b1" & txtLoad.Text & vbCrLf & "\b\f1" & vbCrLf & "}"
Exit Sub
notFound:
MsgBox "Lyrics not Found!", vbCritical, "Error"
Label3 = "Lyrics not Found..."
End Sub

Private Sub Command3_Click()
tOldT1 = Text1
Text3 = "http://www.azlyrics.com/" & Left(Text1, 1) & "/" & Replace(Text1, " ", "") & ".html"
    On Error Resume Next
    
    Dim txt As String
    Dim b() As Byte
    Dim theStart
    Dim theLength
    
    Command3.Enabled = False
    
    b() = Inet2.OpenURL(Text3.Text, 1)
    
    txt = ""
    


    For t = 0 To UBound(b) - 1
        txt = txt + Chr(b(t))
    Next
    
    txtList.Text = txt
    txtList.Text = Replace(txtList.Text, " target=" + Quot + "_blank" + Quot + "", "")
    theStart = InStr(1, txtList.Text, "<FONT Face=Verdana size=5><BR>", vbTextCompare)
theLength = Len(txtList) - theStart
txtList.Text = Mid(txtList, theStart, theLength)
txtList.Text = Replace(txtList.Text, "<FONT Face=Verdana size=5><BR>", "<FONT Face=Verdana size=5>")
theStart = InStr(1, txtList.Text, "</TD></TR>", vbTextCompare)
theLength = Len(txtList) - theStart
txtList.Text = Left(txtList.Text, theStart)
 ' save file
    Dim f As Integer
    f = FreeFile
    Kill "C:\templyrics000.html"
    Open "C:\templyrics000.html" For Binary As #f
    Put #f, , txtList.Text
    Close #f
Command3.Enabled = True
' end of that
WebBrowser1.Navigate "C:\templyrics000.html"
    Exit Sub
noFind:
MsgBox "Cannot find Artist", vbCritical, "Error"
Command3.Enabled = True
End Sub

Private Sub Text5_Change()

End Sub

Private Sub copy_Click()
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "about:<font face=arial size=2>Ready...</font>"
menu.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill "C:\templyrics000.html"
End Sub

Private Sub save_Click()
cd.Filter = "All Files(*.*)|*.*|Rich Text(*.rtf)|*.rtf"
cd.FilterIndex = 2
cd.ShowSave
If cd.FileName <> "" Then
txtLyrics.SaveFile cd.FileName
End If
End Sub

Private Sub txtLyrics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then 'if they right click, 1=left, 2=right
    Form1.PopupMenu menu 'show popup menu
    Else
    DoEvents
End If
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
On Error Resume Next
Dim LyricName
If Left(URL, 9) = "C:\lyrics" Then
If Command1.Enabled = True Then
LyricName = Replace(URL, "C:\lyrics\", "")
Text1 = tOldT1
Text2 = Replace(LyricName, "\", "")
Text2 = Replace(Text2, Replace(Text1, " ", ""), "")
Text2 = Replace(Text2, " ", "")
Text2 = Replace(Text2, ".html", "")
Command1_Click
Cancel = True
Else
MsgBox "Last search not complete", vbInformation, "Notice"
Cancel = True
End If
End If
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
End Sub

