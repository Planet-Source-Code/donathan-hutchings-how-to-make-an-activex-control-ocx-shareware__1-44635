VERSION 5.00
Begin VB.UserControl Text 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtData 
      Height          =   2340
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   15
      Width           =   3015
   End
End
Attribute VB_Name = "Text"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_UserName = ""
Const m_def_SerialNumber = ""

'This constant is used as a return verification for the license file.
Const USER_VERIFIED = 92479252

'Event Declarations:
Event Click() 'MappingInfo=txtData,txtData,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=txtData,txtData,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtData,txtData,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtData,txtData,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtData,txtData,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtData,txtData,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
'Property Variables:
Dim m_UserName As String
Dim m_SerialNumber As String

Private Sub UserControl_Paint()
    On Local Error Resume Next
    txtData.Move 0, 0, Width, Height
    On Local Error GoTo 0
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    txtData.Move 0, 0, Width, Height
    On Local Error GoTo 0
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = txtData.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtData.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = txtData.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtData.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = txtData.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtData.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = txtData.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtData.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = txtData.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    txtData.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    txtData.Refresh
End Sub

Private Sub txtData_Click()
    RaiseEvent Click
End Sub

Private Sub txtData_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub txtData_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub txtData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = txtData.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtData.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = txtData.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtData.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtData.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtData.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtData.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtData.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtData.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtData.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MaxLength
Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
    MaxLength = txtData.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtData.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtData,txtData,-1,MultiLine
Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Returns/sets a value that determines whether a control can accept multiple lines of text."
    MultiLine = txtData.MultiLine
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_UserName = m_def_UserName
    m_SerialNumber = m_def_SerialNumber
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    txtData.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtData.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtData.Enabled = PropBag.ReadProperty("Enabled", True)
    Set txtData.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtData.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtData.Text = PropBag.ReadProperty("Text", "")
    txtData.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    txtData.SelText = PropBag.ReadProperty("SelText", "")
    txtData.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtData.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtData.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    
    'Check to see if the user name and serial number were stored in the property
    'bag.
    m_UserName = PropBag.ReadProperty("UserName", m_def_UserName)
    m_SerialNumber = PropBag.ReadProperty("SerialNumber", m_def_SerialNumber)
    
    'Verify the user name and serial number.
    Select Case VerifyRegistration(m_UserName, m_SerialNumber)
    Case USER_VERIFIED
        Exit Sub
    Case Else
        'Show the about form if not registered.
        frmAbout.Show vbModal, Me
    End Select

    'Use this function to generate the license file.
    'Register
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", txtData.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", txtData.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", txtData.Enabled, True)
    Call PropBag.WriteProperty("Font", txtData.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", txtData.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", txtData.Text, "")
    Call PropBag.WriteProperty("ToolTipText", txtData.ToolTipText, "")
    Call PropBag.WriteProperty("SelText", txtData.SelText, "")
    Call PropBag.WriteProperty("SelStart", txtData.SelStart, 0)
    Call PropBag.WriteProperty("SelLength", txtData.SelLength, 0)
    Call PropBag.WriteProperty("MaxLength", txtData.MaxLength, 0)
    
    'Save the user anem and serial number to the property bag.
    Call PropBag.WriteProperty("UserName", m_UserName, m_def_UserName)
    Call PropBag.WriteProperty("SerialNumber", m_SerialNumber, m_def_SerialNumber)
End Sub

Private Function VerifyRegistration(UserName As String, SerialNumber As String) As Long
    Dim FF As Integer
    Dim Buffer() As Byte
    Dim sTemp As String
    
    On Local Error GoTo ReportError
    
    'Verify the user name and serial number
    Select Case UCase(UserName)
    Case "USER"
        Select Case UCase(SerialNumber)
        Case "SERIAL_NUMBER"
            VerifyRegistration = USER_VERIFIED
            Exit Function
        End Select
    End Select
    
    'If not verified from property bag then get it from the locense file.
    FF = FreeFile
    ReDim Buffer(FileLen(App.Path & "\licfile.lic"))
    Open App.Path & "\licfile.lic" For Binary Access Read As #FF
    Get #FF, , Buffer()
    Close #FF

    sTemp = StrConv(Buffer(), vbUnicode)
    
    'Trim off the null.
    sTemp = Mid(sTemp, 1, Len(sTemp) - 1)
    
    If Len(Trim(sTemp)) > 0 Then
        If InStr(sTemp, "|") > 0 Then
            If UCase(Mid(sTemp, 1, InStr(sTemp, "|") - 1)) = "USER" And _
               UCase(Mid(sTemp, InStr(sTemp, "|") + 1)) = "SERIAL_NUMBER" Then
                    UserName = UCase(Mid(sTemp, 1, InStr(sTemp, "|") - 1))
                    SerialNumber = UCase(Mid(sTemp, InStr(sTemp, "|") + 1))
                    VerifyRegistration = USER_VERIFIED
                    Exit Function
            End If
        End If
    End If
    
ReportError:
    VerifyRegistration = 0
    UserName = ""
    SerialNumber = ""
End Function

Public Sub Register()
    Dim FF As Integer
    Dim Buffer() As Byte
    Dim sTemp As String
    
    sTemp = "USER|SERIAL_NUMBER"
    Buffer() = StrConv(sTemp, vbFromUnicode)
    
    FF = FreeFile
    Open App.Path & "\licfile.lic" For Binary Access Write As #FF
    Put #FF, , Buffer()
    Close #FF

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get UserName() As String
    UserName = m_UserName
End Property

Public Property Let UserName(ByVal New_UserName As String)
    m_UserName = New_UserName
    PropertyChanged "UserName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get SerialNumber() As String
    SerialNumber = m_SerialNumber
End Property

Public Property Let SerialNumber(ByVal New_SerialNumber As String)
    m_SerialNumber = New_SerialNumber
    PropertyChanged "SerialNumber"
End Property

