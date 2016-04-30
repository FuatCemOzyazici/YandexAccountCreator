VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   Caption         =   "Tech Yandex Mail Account Creator"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   3
      Text            =   "Ahmet Yilmaz"
      Top             =   240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1500
      ItemData        =   "Form1.frx":0000
      Left            =   6840
      List            =   "Form1.frx":0002
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Signup"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      ExtentX         =   5530
      ExtentY         =   4048
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
   Begin VB.Label Label4 
      Caption         =   "Ready for action."
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim ayir() As String
ayir = Split(Text1.Text, " ")
WebBrowser1.Document.getelementbyid("firstname").Value = ayir(0)
WebBrowser1.Document.getelementbyid("lastname").Value = ayir(UBound(ayir))
WebBrowser1.Document.getelementbyid("login").Value = Text2.Text
WebBrowser1.Document.getelementbyid("password").Value = Text3.Text
WebBrowser1.Document.getelementbyid("password_confirm").Value = Text3.Text
Dim abc As Object
For Each abc In WebBrowser1.Document.getelementsbytagname("option")
If abc.Value = "18" Then abc.Selected = True
Next abc
WebBrowser1.Document.getelementbyid("hint_answer").Value = "League of Legends"
WebBrowser1.Document.Forms(0).Submit
End Sub

Private Sub Command2_Click()
Randomize
Dim random1 As Long
Dim random2 As Long
Dim random4 As Long
random1 = Int((Rnd * 89) + 10)
random4 = Int((Rnd * 9999) + 1)
random2 = Int((Rnd * List1.ListCount - 1) + 1)
Text2.Text = Replace(List1.List(random2) & Str(random4), " ", "")
Text3.Text = RandomString(6)
Label4.Caption = "Account information generatd. Press the signup button."
End Sub

Private Sub Form_Load()
WebBrowser1.Silent = True
WebBrowser1.Navigate2 "https://passport.yandex.com.tr/registration/mail"

Dim ff As Long
Dim line As String
ff = FreeFile
Open App.Path & "\usernamewordlist.txt" For Input As #ff
Do While Not EOF(ff)
       Line Input #ff, line
       If Len(line) Then List1.AddItem line
Loop
Close #ff
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
 WebBrowser1.Document.body.Scroll = "no"
If InStr(URL, "https://passport.yandex.com.tr/registration/mail") > 0 Then
Dim asd As Object
For Each asd In WebBrowser1.Document.getelementsbytagname("label")
If asd.className = "human-confirmation-switch human-confirmation-via-captcha" Then
asd.Click
End If
Next asd
WebBrowser1.Document.ParentWindow.scrollBy 0, 616
End If

If InStr(URL, "https://passport.yandex.com.tr/passport?mode=passport") > 0 Then

Open App.Path & "\createdaccounts.txt" For Append As #1
Print #1, "Name: " & Text1.Text & " | Username: " & Text2.Text & " | Password: " & Text3.Text & " | Secret Question: Favorite game - League of Legends"
Close #1

Open App.Path & "\createdaccountsmaillist.txt" For Append As #1
Print #1, Text2.Text & "@yandex.com"
Close #1

Label4.Caption = "Account info has been saved to createdaccounts.txt"
Pause 2
Dim aaa As Object
For Each aaa In WebBrowser1.Document.getelementsbytagname("a")
If aaa.innertext = "Çýkýþ" Then aaa.Click
Pause 2
WebBrowser1.Navigate2 "https://passport.yandex.com.tr/registration/mail"
Next aaa
End If


End Sub

