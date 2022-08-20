Option Strict Off
Option Explicit On
Friend Class Form1
	Inherits System.Windows.Forms.Form
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		'Test1 for repeating code
		MYTB.MoveFirst()
		Do Until MYTB.EOF
			If MYTB.Fields("Code").Value = Text1.Text Then
				MsgBox("Error...Enter Code Again")
				Text1.Text = "" : Text1.Focus()
				Exit Sub
			End If
			MYTB.MoveNext()
		Loop 
		'End of test1
		MYTB.AddNew()
		MYTB.Fields("Code").Value = Text1.Text
		MYTB.Fields("Name").Value = Text2.Text
		MYTB.Fields("region").Value = Combo1.Text
		If Option1.Checked = True Then
			MYTB.Fields("Sex").Value = "Male"
		Else
			MYTB.Fields("Sex").Value = "Fema"
		End If
		If Check1.CheckState = 1 Then
			MYTB.Fields("pay1").Value = "yes"
		Else
			MYTB.Fields("pay1").Value = "no"
		End If
		If Check2.CheckState = 1 Then
			MYTB.Fields("pay2").Value = "Yes"
		Else
			MYTB.Fields("pay2").Value = "No"
		End If
		MYTB.Fields("Mobile").Value = Text3.Text
		MYTB.Fields("Bdate").Value = VB6.Format(Text4.Text, "yyyy-mm-dd")
		MYTB.Update()
		MsgBox("Add...Ok")
		'Rem clear
		Text1.Text = "" : Text2.Text = "" : Text3.Text = "" : Text4.Text = ""
		Option1.Checked = False : Option2.Checked = False
		'Check1.Value = 0: Check2.Value = 0
		Check1.CheckState = System.Windows.Forms.CheckState.Unchecked : Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Combo1.Text = ""
		Text1.Focus()
	End Sub
	Private Sub Command10_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command10.Click
		On Error Resume Next
		MYTB.MoveLast()
		Text1.Text = MYTB.Fields("Code").Value
		Text2.Text = MYTB.Fields("Name").Value
		Combo1.Text = MYTB.Fields("region").Value
		Text4.Text = MYTB.Fields("Bdate").Value
		Text3.Text = MYTB.Fields("Mobile").Value
		If MYTB.Fields("Sex").Value = "Male" Then
			Option1.Checked = True
		Else
			Option2.Checked = True
		End If
		If MYTB.Fields("pay1").Value = "Yes" Then
			Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If MYTB.Fields("pay2").Value = "Yes" Then
			Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
	End Sub
	Private Sub Command11_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command11.Click
		End
	End Sub
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
		MYTB.Edit()
		MYTB.Fields("Code").Value = Text1.Text
		MYTB.Fields("Name").Value = Text2.Text
		MYTB.Fields("region").Value = Combo1.Text
		If Option1.Checked = True Then
			MYTB.Fields("Sex").Value = "Male"
		Else
			MYTB.Fields("Sex").Value = "Fema"
		End If
		If Check1.CheckState = 1 Then
			MYTB.Fields("pay1").Value = "Yes"
		Else
			MYTB.Fields("pay1").Value = "No"
		End If
		If Check2.CheckState = 1 Then
			MYTB.Fields("pay2").Value = "Yes"
		Else
			MYTB.Fields("pay2").Value = "No"
		End If
		MYTB.Fields("Mobile").Value = Text3.Text
		MYTB.Fields("Bdate").Value = VB6.Format(Text4.Text, "yyyy-mm-dd")
		MYTB.Update()
		MsgBox("Edit...Ok")
		'Rem clear
		Text1.Text = "" : Text2.Text = "" : Text3.Text = "" : Text4.Text = ""
		Option1.Checked = False : Option2.Checked = False
		'Check1.Value = 0: Check2.Value = 0
		Check1.CheckState = System.Windows.Forms.CheckState.Unchecked : Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Combo1.Text = ""
		Text1.Focus()
	End Sub
	Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
		Dim del As String
		del = CStr(MsgBox("Delet Suore...", MsgBoxStyle.YesNo))
		If del = CStr(MsgBoxResult.Yes) Then
			MYTB.Delete()
			MsgBox("Del...Ok")
		Else
		End If
		'Rem Clear
		Text1.Text = "" : Text2.Text = "" : Text3.Text = "" : Text4.Text = ""
		Option1.Checked = False : Option2.Checked = False
		'Check1.Value = 0: Check2.Value = 0
		Check1.CheckState = System.Windows.Forms.CheckState.Unchecked : Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		Combo1.Text = ""
		Text1.Focus()
	End Sub
	Private Sub Command4_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command4.Click
		'Code
		Dim X As Short
		X = CShort(InputBox("Enter Code"))
		MYTB.MoveFirst()
		Do Until MYTB.EOF
			If MYTB.Fields("Code").Value = X Then
				Text1.Text = MYTB.Fields("Code").Value
				Text2.Text = MYTB.Fields("Name").Value
				Combo1.Text = MYTB.Fields("region").Value
				Text4.Text = MYTB.Fields("Bdate").Value
				Text3.Text = MYTB.Fields("Mobile").Value
				If MYTB.Fields("Sex").Value = "Male" Then
					Option1.Checked = True
				Else
					Option2.Checked = True
				End If
				If MYTB.Fields("pay1").Value = "Yes" Then
					Check1.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				If MYTB.Fields("pay2").Value = "Yes" Then
					Check2.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				Exit Sub
			End If
			MYTB.MoveNext()
		Loop 
		Text1.Focus()
		MsgBox("Error....Code Not Found")
		
	End Sub
	Private Sub Command5_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command5.Click
		'Name
		Dim N As String
		N = InputBox("Enter Name")
		MYTB.MoveFirst()
		Do Until MYTB.EOF
			If MYTB.Fields("Name").Value = N Then
				Text1.Text = MYTB.Fields("Code").Value
				Text2.Text = MYTB.Fields("Name").Value
				Combo1.Text = MYTB.Fields("region").Value
				Text4.Text = MYTB.Fields("Bdate").Value
				Text3.Text = MYTB.Fields("Mobile").Value
				If MYTB.Fields("Sex").Value = "Male" Then
					Option1.Checked = True
				Else
					Option2.Checked = True
				End If
				If MYTB.Fields("pay1").Value = "Yes" Then
					Check1.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				If MYTB.Fields("pay2").Value = "Yes" Then
					Check2.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				Exit Sub
			End If
			MYTB.MoveNext()
		Loop 
		Text1.Focus()
		MsgBox("Error....Name Not Found")
	End Sub
	Private Sub Command6_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command6.Click
		'Mobile
		Dim X As String
		X = InputBox("Enter Name")
		MYTB.MoveFirst()
		Do Until MYTB.EOF
			If MYTB.Fields("Mobile").Value = X Then
				Text1.Text = MYTB.Fields("Code").Value
				Text2.Text = MYTB.Fields("Name").Value
				Combo1.Text = MYTB.Fields("region").Value
				Text4.Text = MYTB.Fields("Bdate").Value
				Text3.Text = MYTB.Fields("Mobile").Value
				If MYTB.Fields("Sex").Value = "Male" Then
					Option1.Checked = True
				Else
					Option2.Checked = True
				End If
				If MYTB.Fields("pay1").Value = "Yes" Then
					Check1.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				If MYTB.Fields("pay2").Value = "Yes" Then
					Check2.CheckState = System.Windows.Forms.CheckState.Checked
				Else
					Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
				End If
				Exit Sub
			End If
			MYTB.MoveNext()
		Loop 
		Text1.Focus()
		MsgBox("Error....Mobile Not Found")
	End Sub
	
	Private Sub Command7_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command7.Click
		On Error Resume Next
		MYTB.MoveFirst()
		Text1.Text = MYTB.Fields("Code").Value
		Text2.Text = MYTB.Fields("Name").Value
		Combo1.Text = MYTB.Fields("region").Value
		Text4.Text = MYTB.Fields("Bdate").Value
		Text3.Text = MYTB.Fields("Mobile").Value
		If MYTB.Fields("Sex").Value = "Male" Then
			Option1.Checked = True
		Else
			Option2.Checked = True
		End If
		If MYTB.Fields("pay1").Value = "Yes" Then
			Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If MYTB.Fields("pay2").Value = "Yes" Then
			Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
	End Sub
	
	Private Sub Command8_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command8.Click
		On Error Resume Next
		MYTB.MoveNext()
		Text1.Text = MYTB.Fields("Code").Value
		Text2.Text = MYTB.Fields("Name").Value
		Combo1.Text = MYTB.Fields("region").Value
		Text4.Text = MYTB.Fields("Bdate").Value
		Text3.Text = MYTB.Fields("Mobile").Value
		If MYTB.Fields("Sex").Value = "Male" Then
			Option1.Checked = True
		Else
			Option2.Checked = True
		End If
		If MYTB.Fields("pay1").Value = "Yes" Then
			Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If MYTB.Fields("pay2").Value = "Yes" Then
			Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	Private Sub Command9_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command9.Click
		On Error Resume Next
		MYTB.MovePrevious()
		Text1.Text = MYTB.Fields("Code").Value
		Text2.Text = MYTB.Fields("Name").Value
		Combo1.Text = MYTB.Fields("region").Value
		Text4.Text = MYTB.Fields("Bdate").Value
		Text3.Text = MYTB.Fields("Mobile").Value
		If MYTB.Fields("Sex").Value = "Male" Then
			Option1.Checked = True
		Else
			Option2.Checked = True
		End If
		If MYTB.Fields("pay1").Value = "Yes" Then
			Check1.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check1.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		If MYTB.Fields("pay2").Value = "Yes" Then
			Check2.CheckState = System.Windows.Forms.CheckState.Checked
		Else
			Check2.CheckState = System.Windows.Forms.CheckState.Unchecked
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Form event Form1.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub Form1_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		'Form1.BackColor = vbyellow
		Me.BackColor = System.Drawing.ColorTranslator.FromOle(&HC0FFFF)
		Command11.BackColor = System.Drawing.Color.Red
	End Sub
	Private Sub Form1_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		MYDB = DAODBEngine_definst.OpenDatabase("D:\Education 2020\Students.mdb")
		MYTB = MYDB.OpenRecordset("student")
		
	End Sub
	Private Sub Text1_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles Text1.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = 13 Then Text2.Focus()
	End Sub
	Private Sub Text2_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles Text2.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = 13 Then Combo1.Focus()
	End Sub
	Private Sub Text4_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles Text4.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		'..Test2 for birth date
		If KeyCode = 13 Then
			If IsDate(Text4.Text) = False Then
				MsgBox("Error....Enter Date Again")
				Text4.Text = "" : Text4.Focus()
				Exit Sub
			End If
			Text4.Text = VB6.Format(Text4.Text, "yyyy-mm-dd")
			Text4.Focus()
		End If
		'End of test2
	End Sub
	Private Sub text3_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles text3.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		If KeyCode = 13 Then Text4.Focus()
	End Sub
End Class