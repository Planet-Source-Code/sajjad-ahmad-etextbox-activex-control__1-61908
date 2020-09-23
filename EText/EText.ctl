VERSION 5.00
Begin VB.UserControl TextBox 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1710
   PropertyPages   =   "EText.ctx":0000
   ScaleHeight     =   300
   ScaleWidth      =   1710
   ToolboxBitmap   =   "EText.ctx":002D
   Begin VB.TextBox ETextBox 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Const EM_LINELENGTH = &HC1

Public Enum CharacterType
    Normal
    Letters
    Numbers
    LettersAndNumbers
    Money
    Percent
    Fraction
    Decimals
    Dates
End Enum

Public Enum CaseType
    Normal
    UpperCase
    LowerCase
End Enum

Enum uAlignment
    Left_Justify = 0
    Right_Justify = 1
    Center = 2
End Enum

Enum uBorderStyle
    None = 0
    Fixed_Single = 1
End Enum

Enum uAppearance
    Appear_Flat = 0
    Appear_3D = 1
End Enum
'Default Property Values:
Const m_def_TextType = 0
Const m_def_TextCaseType = 0
'Property Variables:
Dim m_TextType As CharacterType
Dim m_TextCaseType As CaseType
'Event Declarations:
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=ETextBox,ETextBox,-1,KeyUp
Event KeyPress(KeyAscii As Integer) 'MappingInfo=ETextBox,ETextBox,-1,KeyPress
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=ETextBox,ETextBox,-1,KeyDown
Event Change() 'MappingInfo=ETextBox,ETextBox,-1,Change
Event Click() 'MappingInfo=ETextBox,ETextBox,-1,Click
Event DblClick() 'MappingInfo=ETextBox,ETextBox,-1,DblClick

Private Function InsertStr(ByVal InsertTo As String, ByVal Str As String, ByVal Position As Integer) As String
    Dim Str1 As String
    Dim Str2 As String
    
    Str1 = Mid(InsertTo, 1, Position - 1)
    Str2 = Mid(InsertTo, Position, Len(InsertTo) - Len(Str1))
    
    InsertStr = Str1 & Str & Str2
End Function

Private Sub ETextBox_GotFocus()
    ETextBox.BackColor = &H80000018
    ETextBox.SelStart = 0
    ETextBox.SelLength = Len(ETextBox.Text)
End Sub

Private Sub ETextBox_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
    Dim lCurPos As Long
    Dim lLineLength As Long
    Dim I As Integer
    
    Dim MoneyDolB As Boolean
    Dim MoneyDotB As Boolean
    Dim MoneyDot As String
    Dim MoneyDolLoc As Long
    Dim MoneyDotLoc As Long
    
    Dim PercentDotB As Boolean
    Dim PercentPerB As Boolean
    Dim PercentNum As String
    Dim PercentDot As String
    Dim PercentLoc As Long
    Dim PercentDotLoc As Long
    
    Dim DecimalDotB As Boolean
    
    Dim Space As Boolean
    Dim FractionSlash As Boolean
    Dim SpaceLoc As Long
    Dim FractionLoc As Long
    
    Select Case TextType
    Case Letters
        If Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    Case Numbers
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    Case LettersAndNumbers
        If IsNumeric(Chr(KeyAscii)) = False And (Not (KeyAscii > 64 And KeyAscii < 91) And Not (KeyAscii > 96 And KeyAscii < 123)) And KeyAscii <> 8 And KeyAscii <> 32 Then
            KeyAscii = 0
            Beep
            Exit Sub
        End If
    Case Money
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 36 And KeyAscii <> 46 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If ETextBox.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If ETextBox.SelLength = 0 Then
                lCurPos = ETextBox.SelStart
            Else
                lCurPos = ETextBox.SelStart + ETextBox.SelLength
            End If
            
            
            ' Determine textbox length
            lLineLength = SendMessage(ETextBox.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            
            ' Determine location/existance of "$" and "."
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "$" Then
                    MoneyDolB = True
                    MoneyDolLoc = I
                    Exit For
                End If
            Next I
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "." Then
                    MoneyDotB = True
                    MoneyDotLoc = I
                    Exit For
                End If
            Next I
                        
            ' Make sure number only goes to 2 decimal places
            If MoneyDotB = True Then
                MoneyDot = Mid(ETextBox.Text, InStr(1, ETextBox.Text, ".") + 1, Len(ETextBox.Text) + InStr(1, ETextBox.Text, ".") + 1)
        
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc + 1 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(MoneyDot) = 2 And lCurPos = MoneyDotLoc And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = MoneyDotLoc + 2 And KeyAscii <> 8 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If
                
            ' Make sure "." and "$" is only typed once
            If KeyAscii = 36 And MoneyDolB = False Then
                MoneyDolB = True
            ElseIf KeyAscii = 36 And MoneyDolB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> 0 And MoneyDolB <> False And KeyAscii = 36 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If

            If KeyAscii = 46 And MoneyDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = 46 And MoneyDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    Case Percent
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 37 And KeyAscii <> 46 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If ETextBox.SelLength <> 0 Then
                Exit Sub
            End If
            
            ' Determine cursor position
            If ETextBox.SelLength = 0 Then
                lCurPos = ETextBox.SelStart
            Else
                lCurPos = ETextBox.SelStart + ETextBox.SelLength
            End If
            
            
            ' Determine textbox length
            lLineLength = SendMessage(ETextBox.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            
            ' Determine location of "%" and "."
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "%" Then
                    PercentPerB = True
                    PercentLoc = I
                    Exit For
                End If
            Next I
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "." Then
                    PercentDotB = True
                    PercentDotLoc = I
                    Exit For
                End If
            Next I


            ' Make sure number only goes to 2 decimal places
            If PercentDotB = True Then
                PercentDot = Mid(ETextBox.Text, InStr(1, ETextBox.Text, ".") + 1, Len(ETextBox.Text) + InStr(1, ETextBox.Text, ".") + 1)
        
                If InStr(1, PercentDot, "%") <> 0 Then
                    PercentDot = Mid(PercentDot, 1, Len(PercentDot) - 1)
                End If
        
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc + 1 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If Len(PercentDot) = 2 And lCurPos = PercentDotLoc And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
                If lCurPos = PercentDotLoc + 2 And KeyAscii <> 8 And KeyAscii <> 37 Then
                    KeyAscii = 0
                    Beep
                    Exit Sub
                End If
            End If


            ' Make sure "%" and "." is only typed once
            If KeyAscii = 37 And PercentPerB = False Then
                PercentPerB = True
            ElseIf KeyAscii = 37 And PercentPerB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            If lCurPos <> Len(ETextBox.Text) And PercentPerB <> False And KeyAscii = 37 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If KeyAscii = 46 And PercentDotB = False Then
                MoneyDotB = True
            ElseIf KeyAscii = 46 And PercentDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            
            ' Make sure numbers are not written after the "%"
            If KeyAscii <> 37 And KeyAscii <> 8 And PercentPerB = True And lCurPos = PercentLoc Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            ' Determine if the percentage is >100
            If IsNumeric(Chr(KeyAscii)) = True Then
                PercentNum = ETextBox.Text
                PercentNum = InsertStr(PercentNum, Chr(KeyAscii), lCurPos + 1)
                If InStr(1, PercentNum, "%") <> 0 Then
                    If Val(Mid(PercentNum, 1, Len(PercentNum) - 1)) > 100 Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                Else
                    If Val(PercentNum) > 100 Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                End If
            End If
        End If
    Case Fraction
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 47 And KeyAscii <> 32 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            If ETextBox.SelLength <> 0 Then
                Exit Sub
            End If
            ' Determine cursor position
            If ETextBox.SelLength = 0 Then
                lCurPos = ETextBox.SelStart
            Else
                lCurPos = ETextBox.SelStart + ETextBox.SelLength
            End If
            
            
            ' Determine textbox length
            lLineLength = SendMessage(ETextBox.hwnd, EM_LINELENGTH, lCurPos, 0)
            
            
            ' Determine location of " " and "/"
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "/" Then
                    FractionLoc = I
                    Exit For
                End If
            Next I
    
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = " " Then
                    SpaceLoc = I
                    Exit For
                End If
            Next I
            
            If FractionLoc <> 0 Then
                FractionSlash = True
            End If
            If SpaceLoc <> 0 Then
                Space = True
            End If
            
            
            ' Don't allow more then 1 space in the field
            If (Space = True Or Fraction = True) And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If Space = False And KeyAscii = 32 Then
                Space = True
            End If
            
    
            ' Check if " " is being used correctly
            If lCurPos = 0 And KeyAscii = 32 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            
            ' Don't allow more then 1 "/" in the field
            If FractionSlash = True And KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
            
            If FractionSlash = False And KeyAscii = 47 Then
                FractionSlash = True
            End If
            
            
            ' Check if "/" is being used correctly
            If lLineLength >= 1 Then
                If lCurPos > 0 Then
                    If KeyAscii = 47 And IsNumeric(Mid(ETextBox.Text, lCurPos, 1)) = False Then
                        KeyAscii = 0
                        Beep
                        Exit Sub
                    End If
                End If
            ElseIf KeyAscii = 47 Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    Case Decimals
        If IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
            Exit Sub
        Else
            ' Determine textbox length
            lLineLength = SendMessage(ETextBox.hwnd, EM_LINELENGTH, lCurPos, 0)
        
            ' Determine existance of "."
            For I = 1 To lLineLength
                If Mid(ETextBox.Text, I, 1) = "." Then
                    DecimalDotB = True
                    Exit For
                End If
            Next I
                        
                
            ' Make sure "." is only typed once
            If KeyAscii = 46 And DecimalDotB = False Then
                DecimalDotB = True
            ElseIf KeyAscii = 46 And DecimalDotB = True Then
                KeyAscii = 0
                Beep
                Exit Sub
            End If
        End If
    End Select
    
    Select Case TextCaseType
    Case UpperCase
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Case LowerCase
        KeyAscii = Asc(LCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub ETextBox_LostFocus()
    
    ETextBox.BackColor = vbWhite
    
    If TextType = Money Then
        If ETextBox.Text = "" Then
            ETextBox.Text = "0.00"
        End If
        'If InStr(1, ETextBox.Text, "$") = 0 Then
        '    ETextBox.Text = "$" & ETextBox.Text
        'End If
        If InStr(1, ETextBox.Text, ".") = 0 Then
            ETextBox.Text = ETextBox.Text & ".00"
        End If
        If InStr(1, ETextBox.Text, ".") <> 0 Then
            If Len(ETextBox.Text) <> InStr(1, ETextBox.Text, ".") + 2 Then
                If Len(ETextBox.Text) = InStr(1, ETextBox.Text, ".") + 1 Then
                    ETextBox.Text = ETextBox.Text & "0"
                Else
                    ETextBox.Text = ETextBox.Text & "00"
                End If
            End If
            If Mid(ETextBox.Text, 2, 1) = "." Then
                ETextBox.Text = "0" & Mid(ETextBox.Text, 2)
            End If
        End If
    ElseIf TextType = Percent Then
        If ETextBox.Text = "" Then
            ETextBox.Text = "0%"
        End If
        If InStr(1, ETextBox.Text, "%") = 0 Then
            ETextBox.Text = ETextBox.Text & "%"
        End If
        If InStr(1, ETextBox.Text, "%") <> 0 Then
            If Len(ETextBox.Text) = 1 Then
                ETextBox.Text = "0%"
            End If
        End If
        If InStr(1, ETextBox.Text, ".") <> 0 Then
            If Mid(ETextBox.Text, 1, Len(ETextBox.Text) - 1) = "." Then
                ETextBox.Text = Mid(ETextBox.Text, 1, Len(ETextBox.Text) - 2) & "%"
            End If
            If Mid(ETextBox.Text, 1, 1) = "." Then
                ETextBox.Text = "0" & Mid(ETextBox.Text, 1)
            End If
        End If
    ElseIf TextType = Numbers And ETextBox.Text = "" Then
        ETextBox.Text = "0"
    ElseIf TextType = Fraction Then
        If ETextBox.Text = "" Then
            ETextBox.Text = "0"
        End If
        
        ' if the user inputs a fractional number
        If InStr(1, ETextBox.Text, "/") <> 0 Then
            ' if / is the first character in the text box then set to 0
            If InStr(1, ETextBox.Text, "/") = 1 Then
                ETextBox.Text = "0"
            ' make sure there are numbers before and after the /
            ElseIf (IsNumeric(Mid(ETextBox.Text, InStr(1, ETextBox.Text, "/") - 1, 1)) = False) Or (IsNumeric(Mid(ETextBox.Text, InStr(1, ETextBox.Text, "/") + 1, 1)) = False) Then
                ETextBox.Text = "0"
            End If
        End If
        ETextBox.Text = Trim(ETextBox.Text)
        
    ElseIf TextType = Decimals And ETextBox.Text = "" Then
        ETextBox.Text = "0"
        
    ElseIf TextType = Dates Then
        If ETextBox.Text = "" Or ETextBox.Text = "00/00/0000" Then Exit Sub
        If Not IsDate(ETextBox.Text) Then
            Beep
            ETextBox.ForeColor = vbRed
            ETextBox.SetFocus
            Exit Sub
        Else
            ETextBox.Text = Format(ETextBox, "Short Date")
            ETextBox.ForeColor = vbBlack
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    With ETextBox
        .Height = UserControl.ScaleHeight
        .Top = UserControl.ScaleTop
        .Left = UserControl.ScaleLeft
        .Width = UserControl.ScaleWidth
    End With
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_TextType = m_def_TextType
    m_TextCaseType = m_def_TextCaseType
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ETextBox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set ETextBox.Font = PropBag.ReadProperty("Font", Ambient.Font)
    ETextBox.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    ETextBox.Text = PropBag.ReadProperty("Text", "")
    ETextBox.PasswordChar = PropBag.ReadProperty("PasswordChar", "")
    ETextBox.Enabled = PropBag.ReadProperty("Enabled", True)
    m_TextType = PropBag.ReadProperty("TextType", m_def_TextType)
    m_TextCaseType = PropBag.ReadProperty("TextCaseType", m_def_TextCaseType)
    ETextBox.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    ETextBox.Locked = PropBag.ReadProperty("Locked", False)
    ETextBox.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    ETextBox.Text = PropBag.ReadProperty("Display", "")
    ETextBox.Alignment = PropBag.ReadProperty("Alignment", 0)
    ETextBox.Appearance = PropBag.ReadProperty("Appearance", 1)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", ETextBox.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", ETextBox.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", ETextBox.BorderStyle, 1)
    Call PropBag.WriteProperty("Text", ETextBox.Text, "")
    Call PropBag.WriteProperty("PasswordChar", ETextBox.PasswordChar, "")
    Call PropBag.WriteProperty("Enabled", ETextBox.Enabled, True)
    Call PropBag.WriteProperty("TextType", m_TextType, m_def_TextType)
    Call PropBag.WriteProperty("TextCaseType", m_TextCaseType, m_def_TextCaseType)
    Call PropBag.WriteProperty("MaxLength", ETextBox.MaxLength, 0)
    Call PropBag.WriteProperty("Locked", ETextBox.Locked, False)
    Call PropBag.WriteProperty("ForeColor", ETextBox.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Display", ETextBox.Text, "")
    Call PropBag.WriteProperty("Alignment", ETextBox.Alignment, 0)
    Call PropBag.WriteProperty("Appearance", ETextBox.Appearance, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = ETextBox.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    ETextBox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Font
Public Property Get Font() As Font
    Set Font = ETextBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set ETextBox.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,BorderStyle
Public Property Get BorderStyle() As uBorderStyle
    BorderStyle = ETextBox.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As uBorderStyle)
    ETextBox.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Refresh
Public Sub Refresh()
    ETextBox.Refresh
End Sub

Private Sub ETextBox_Click()
    RaiseEvent Click
End Sub

Private Sub ETextBox_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Text
Public Property Get Text() As String
    Text = ETextBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    ETextBox.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,PasswordChar
Public Property Get PasswordChar() As String
    PasswordChar = ETextBox.PasswordChar
End Property

Public Property Let PasswordChar(ByVal New_PasswordChar As String)
    ETextBox.PasswordChar() = New_PasswordChar
    PropertyChanged "PasswordChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = ETextBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    ETextBox.Enabled() = New_Enabled
    
    If ETextBox.Enabled = True Then
        ETextBox.BackColor = &HFFFFFF
    ElseIf ETextBox.Enabled = False Then
        ETextBox.BackColor = &HC0C0C0
    End If

    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,MultiLine
Public Property Get MultiLine() As Boolean
    MultiLine = ETextBox.MultiLine
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=23,0,0,0
Public Property Get TextType() As CharacterType
    TextType = m_TextType
End Property

Public Property Let TextType(ByVal New_TextType As CharacterType)
    m_TextType = New_TextType

    If New_TextType = Money Then
        ETextBox.Text = "0.00"
        ETextBox.Alignment = 1
    ElseIf New_TextType = Percent Then
        ETextBox.Text = "0%"
        ETextBox.Alignment = 1
    ElseIf New_TextType = Numbers Then
        ETextBox.Text = "0"
        ETextBox.Alignment = 1
    ElseIf New_TextType = Fraction Then
        ETextBox.Text = "0"
        ETextBox.Alignment = 1
    ElseIf New_TextType = Decimals Then
        ETextBox.Text = "0"
        ETextBox.Alignment = 1
    ElseIf New_TextType = Dates Then
        ETextBox.Text = "00/00/0000"
        ETextBox.Alignment = 1
    Else
        ETextBox.Text = ""
    End If
    
    PropertyChanged "TextType"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,0,0,0
Public Property Get TextCaseType() As CaseType
    TextCaseType = m_TextCaseType
End Property

Public Property Let TextCaseType(ByVal New_TextCaseType As CaseType)
    m_TextCaseType = New_TextCaseType
    PropertyChanged "TextCaseType"
End Property

Private Sub ETextBox_Change()
    RaiseEvent Change
End Sub

Private Sub ETextBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub ETextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    Select Case KeyCode
        Case 13 '39, 40, 13  Next Control: right arrow, down arrow and Enter
            SendKeys "{Tab}"
        'Case 37, 38 'Previous Control: left and up arrows
            'SendKeys "+{Tab}"
    End Select

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = ETextBox.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    ETextBox.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Locked
Public Property Get Locked() As Boolean
    Locked = ETextBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    ETextBox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = ETextBox.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    ETextBox.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Text
Public Property Get Display() As String
Attribute Display.VB_ProcData.VB_Invoke_Property = "Display"
Attribute Display.VB_MemberFlags = "24"
    Display = ETextBox.Text
End Property

Public Property Let Display(ByVal New_Display As String)
    ETextBox.Text() = New_Display
    PropertyChanged "Display"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Alignment
Public Property Get Alignment() As uAlignment
    Alignment = ETextBox.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As uAlignment)
    ETextBox.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ETextBox,ETextBox,-1,Appearance
Public Property Get Appearance() As uAppearance
    Appearance = ETextBox.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As uAppearance)
    ETextBox.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property


