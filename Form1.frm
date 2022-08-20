VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   26
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton Command11 
      Caption         =   "EXIT"
      Height          =   855
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "LAST"
      Height          =   495
      Left            =   240
      TabIndex        =   23
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "BACK"
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "FIRST"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TEL"
      Height          =   495
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NAME"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CODE"
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "ﬁ”ÿ2"
      Height          =   495
      Left            =   5760
      TabIndex        =   13
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ﬁ”ÿ1"
      Height          =   495
      Left            =   7920
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "«‰ÀÏ"
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "–ﬂ—"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   9
      Top             =   6360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   7560
      MaxLength       =   12
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5880
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   8160
      MaxLength       =   4
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "  «·„Ì·«œ"
      Height          =   495
      Left            =   9840
      TabIndex        =   25
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ﬂÊœ"
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "«·„Õ„Ê·"
      Height          =   495
      Left            =   9840
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "«·„’—Ê›« "
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "«·‰Ê⁄"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "„‰ÿﬁ…"
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "«”„"
      Height          =   495
      Left            =   9840
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Test1 for repeating code
MYTB.MoveFirst
Do Until MYTB.EOF
If MYTB!Code = Text1 Then
MsgBox "Error...Enter Code Again"
Text1 = "": Text1.SetFocus
Exit Sub
End If
MYTB.MoveNext
Loop
'End of test1
MYTB.AddNew
MYTB!Code = Text1
MYTB!Name = Text2
MYTB!region = Combo1
If Option1.Value = True Then
MYTB!Sex = "Male"
Else
MYTB!Sex = "Fema"
End If
If Check1.Value = 1 Then
MYTB!pay1 = "yes"
Else
MYTB!pay1 = "no"
End If
If Check2.Value = 1 Then
MYTB!pay2 = "Yes"
Else
MYTB!pay2 = "No"
End If
MYTB!Mobile = Text3
MYTB!Bdate = Format(Text4, "yyyy-mm-dd")
MYTB.Update
MsgBox "Add...Ok"
'Rem clear
Text1 = "": Text2 = "": Text3 = "": Text4 = ""
Option1.Value = False: Option2.Value = False
'Check1.Value = 0: Check2.Value = 0
Check1 = 0: Check2 = 0
Combo1 = ""
Text1.SetFocus
End Sub
Private Sub Command10_Click()
On Error Resume Next
MYTB.MoveLast
Text1 = MYTB!Code
Text2 = MYTB!Name
Combo1 = MYTB!region
Text4 = MYTB!Bdate
Text3 = MYTB!Mobile
If MYTB!Sex = "Male" Then
Option1.Value = True
Else
Option2.Value = True
End If
If MYTB!pay1 = "Yes" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If MYTB!pay2 = "Yes" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
End Sub
Private Sub Command11_Click()
End
End Sub
Private Sub Command2_Click()
MYTB.Edit
        MYTB!Code = Text1
        MYTB!Name = Text2
        MYTB!region = Combo1
        If Option1.Value = True Then
            MYTB!Sex = "Male"
        Else
            MYTB!Sex = "Fema"
        End If
        If Check1.Value = 1 Then
            MYTB!pay1 = "Yes"
        Else
            MYTB!pay1 = "No"
        End If
        If Check2.Value = 1 Then
            MYTB!pay2 = "Yes"
        Else
            MYTB!pay2 = "No"
        End If
MYTB!Mobile = Text3
        MYTB!Bdate = Format(Text4, "yyyy-mm-dd")
        MYTB.Update
        MsgBox "Edit...Ok"
        'Rem clear
        Text1 = "": Text2 = "": Text3 = "": Text4 = ""
        Option1.Value = False: Option2.Value = False
'Check1.Value = 0: Check2.Value = 0
Check1 = 0: Check2 = 0
        Combo1 = ""
        Text1.SetFocus
End Sub
Private Sub Command3_Click()
Dim del As String
        del = MsgBox("Delet Suore...", vbYesNo)
        If del = vbYes Then
            MYTB.Delete
            MsgBox "Del...Ok"
        Else
        End If
        'Rem Clear
        Text1 = "": Text2 = "": Text3 = "": Text4 = ""
        Option1.Value = False: Option2.Value = False
        'Check1.Value = 0: Check2.Value = 0
        Check1 = 0: Check2 = 0
        Combo1 = ""
        Text1.SetFocus
End Sub
Private Sub Command4_Click()
'Code
        Dim X As Integer
        X = InputBox("Enter Code")
        MYTB.MoveFirst
        Do Until MYTB.EOF
            If MYTB!Code = X Then
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If
                Exit Sub
End If
                MYTB.MoveNext
Loop
        Text1.SetFocus
        MsgBox "Error....Code Not Found"

End Sub
Private Sub Command5_Click()
'Name
        Dim N As String
        N = InputBox("Enter Name")
        MYTB.MoveFirst
        Do Until MYTB.EOF
            If MYTB!Name = N Then
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If
                Exit Sub
End If
                MYTB.MoveNext
        Loop
        Text1.SetFocus
        MsgBox "Error....Name Not Found"
End Sub
Private Sub Command6_Click()
'Mobile
        Dim X As String
        X = InputBox("Enter Name")
        MYTB.MoveFirst
        Do Until MYTB.EOF
            If MYTB!Mobile = X Then
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If
                Exit Sub
End If
                MYTB.MoveNext
        Loop
        Text1.SetFocus
        MsgBox "Error....Mobile Not Found"
End Sub

Private Sub Command7_Click()
On Error Resume Next
MYTB.MoveFirst
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
MYTB.MoveNext
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If

End Sub

Private Sub Command9_Click()
On Error Resume Next
MYTB.MovePrevious
                Text1 = MYTB!Code
                Text2 = MYTB!Name
                Combo1 = MYTB!region
                Text4 = MYTB!Bdate
                Text3 = MYTB!Mobile
                If MYTB!Sex = "Male" Then
                    Option1.Value = True
                Else
                    Option2.Value = True
                End If
                If MYTB!pay1 = "Yes" Then
                    Check1.Value = 1
                Else
                    Check1.Value = 0
                End If
                If MYTB!pay2 = "Yes" Then
                    Check2.Value = 1
                Else
                    Check2.Value = 0
                End If

End Sub

Private Sub Form_Activate()
'Form1.BackColor = vbyellow
        Form1.BackColor = &HC0FFFF
        Command11.BackColor = vbRed
End Sub
Private Sub Form_Load()
Set MYDB = OpenDatabase("D:\Education 2020\Students.mdb")
        Set MYTB = MYDB.OpenRecordset("student")

End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then Text2.SetFocus
    End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then Combo1.SetFocus
    End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
Rem..Test2 for birth date
        If KeyCode = 13 Then
            If IsDate(Text4) = False Then
                MsgBox "Error....Enter Date Again"
                Text4 = "": Text4.SetFocus
                Exit Sub
            End If
            Text4 = Format(Text4, "yyyy-mm-dd")
            Text4.SetFocus
        End If
'End of test2
End Sub
Private Sub text3_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then Text4.SetFocus
    End Sub

