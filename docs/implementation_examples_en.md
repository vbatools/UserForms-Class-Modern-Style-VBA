# Implementation Examples for VBA Modern Style UserForms Project

## Table of Contents
1. [Basic Implementation](#basic-implementation)
2. [Advanced Color Configuration](#advanced-color-configuration)
3. [Creating Toggle Switches](#creating-toggle-switches)
4. [Adding Icons](#adding-icons)
5. [Dynamic Style Configuration](#dynamic-style-configuration)
6. [Working with Style Collection](#working-with-style-collection)
7. [Integration with Existing Forms](#integration-with-existing-forms)
8. [Event Handling Examples](#event-handling-examples)
9. [Creating Themes](#creating-themes)
10. [Advanced Examples](#advanced-examples)

## Basic Implementation

### Simple Usage Example
```vba
' In the form module
Dim MStyleItem As clsModernStyle

Private Sub UserForm_Initialize()
    ' Create an instance of the style class
    Set MStyleItem = New clsModernStyle
    
    ' Initialize styles for the current form
    Call MStyleItem.Initialize(Me)
End Sub
```

### Example with Multiple Controls
```vba
Private Sub UserForm_Initialize()
    ' Add controls to the form through the designer
    ' (TextBox1, ComboBox1, CheckBox1, OptionButton1)
    
    Set MStyleItem = New clsModernStyle
    
    ' Set tooltip texts for controls
    TextBox1.ControlTipText = "Username"
    ComboBox1.ControlTipText = "Select option"
    
    ' Initialize styles
    Call MStyleItem.Initialize(Me)
End Sub
```

## Advanced Color Configuration

### Configuring Colors During Initialization
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Initialize with custom colors
    MStyleItem.Initialize Me, _
        ColorBarTitleOn:=RGB(0, 100, 200), _      ' Active title color
        ColorBarTitleOff:=RGB(120, 120, 120), _    ' Inactive title color
        ColorBarBottomOn:=RGB(0, 100, 200), _      ' Active bottom line color
        ColorBarBottomOff:=RGB(180, 180, 180), _   ' Inactive bottom line color
        ColorBackGroundOn:=RGB(255, 255, 255), _   ' Active background color
        ColorBackGroundOff:=RGB(245, 245, 245), _  ' Inactive background color
        ColorBarIconOn:=RGB(0, 100, 200), _        ' Active icon color
        ColorBarIconOff:=RGB(150, 150, 150), _     ' Inactive icon color
        ColorDropArrowOn:=RGB(0, 100, 200), _      ' Active arrow color
        ColorDropArrowOff:=RGB(150, 150, 150)      ' Inactive arrow color
End Sub
```

### Using Predefined Color Schemes
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Dark theme color scheme
    If ThemeManager.IsDarkTheme Then
        MStyleItem.Initialize Me, _
            ColorBarTitleOn:=RGB(255, 255, 255), _
            ColorBarTitleOff:=RGB(180, 180, 180), _
            ColorBarBottomOn:=RGB(0, 120, 215), _
            ColorBarBottomOff:=RGB(80, 80, 80), _
            ColorBackGroundOn:=RGB(30, 30), _
            ColorBackGroundOff:=RGB(20, 20)
    Else
        ' Light theme color scheme
        MStyleItem.Initialize Me, _
            ColorBarTitleOn:=RGB(0, 0, 0), _
            ColorBarTitleOff:=RGB(120, 120, 120), _
            ColorBarBottomOn:=RGB(0, 100, 200), _
            ColorBarBottomOff:=RGB(180, 180, 180), _
            ColorBackGroundOn:=RGB(255, 255, 255), _
            ColorBackGroundOff:=RGB(245, 245, 245)
    End If
End Sub
```

## Creating Toggle Switches

### Simple Toggle Switch
```vba
Private Sub UserForm_Initialize()
    ' Set Tag property to create a toggle switch
    CheckBox1.Tag = "SWITCH"
    CheckBox2.Tag = "SWITCH"
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Toggle Switch Group
```vba
Private Sub UserForm_Initialize()
    ' Create a group of toggle switches
    ToggleOption1.Caption = "Option 1"
    ToggleOption1.Tag = "SWITCH"
    
    ToggleOption2.Caption = "Option 2"
    ToggleOption2.Tag = "SWITCH"
    
    ToggleOption3.Caption = "Option 3"
    ToggleOption3.Tag = "SWITCH"
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Set initial state
    ToggleOption1.Value = True
End Sub
```

### Dynamic Toggle Creation
```vba
Private Sub UserForm_Initialize()
    ' Dynamically create toggle switches
    Dim chk As MSForms.CheckBox
    Dim i As Integer
    
    For i = 1 To 5
        Set chk = Me.Controls.Add("Forms.CheckBox.1", "DynamicToggle" & i, True)
        With chk
            .Left = 20
            .Top = 50 + (i - 1) * 30
            .Width = 200
            .Height = 20
            .Caption = "Toggle " & i
            .Tag = "SWITCH"
        End With
    Next i
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

## Adding Icons

### Using Built-in Icons
```vba
Private Sub UserForm_Initialize()
    ' Set icons using numeric values from enumIcons
    TextBox1.Tag = 59193  ' CheckBox1
    TextBox2.Tag = 59188  ' FavoriteStar
    TextBox3.Tag = 60241  ' Heart
    
    ComboBox1.Tag = 61735  ' PaginationDotSolid10
    ListBox1.Tag = 59962   ' CircleRing
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Using Icons for Different Control Types
```vba
Private Sub UserForm_Initialize()
    ' Icons for text boxes
    UsernameBox.Tag = 59193   ' Square for username
    PasswordBox.Tag = 149     ' Password character
    EmailBox.Tag = 59188      ' Star for email
    
    ' Icons for combo boxes
    CountryCombo.Tag = 60619  ' Radio button for country
    CategoryCombo.Tag = 61804 ' Square with checkmark for category
    
    ' Icons for checkboxes
    AgreementBox.Tag = 59194  ' Square with checkmark for agreement
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Dynamic Icon Addition
```vba
Private Sub UserForm_Initialize()
    Dim controlsList As Variant
    Dim iconsList As Variant
    Dim i As Integer
    
    ' List of controls and corresponding icons
    controlsList = Array("TextBox1", "TextBox2", "ComboBox1", "CheckBox1")
    iconsList = Array(59193, 59188, 60619, 59194) ' Values from enumIcons
    
    ' Apply icons
    For i = 0 To UBound(controlsList)
        Me.Controls(controlsList(i)).Tag = iconsList(i)
    Next i
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

## Dynamic Style Configuration

### Changing Styles After Initialization
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Change styles of specific elements after initialization
    With MStyleItem.getItemByName(TextBox1.Name)
        .ColorBackGroundOff = RGB(255, 250, 200) ' Yellow background
        .ColorBackGroundOn = RGB(255, 255, 255)  ' White background on focus
        .ColorBarBottomOn = RGB(255, 0, 0)        ' Red line on focus
    End With
End Sub

Private Sub ChangeStyleButton_Click()
    ' Dynamic style change on button click
    With MStyleItem.getItemByName(TextBox2.Name)
        .ColorBarTitleOn = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
        .ColorBarBottomOn = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    End With
End Sub
```

### Conditional Style Configuration
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Conditional style configuration based on data type
    Dim ctrl As MSForms.control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            With MStyleItem.getItemByName(ctrl.Name)
                ' If control name contains "password", set special style
                If InStr(LCase(ctrl.Name), "password") > 0 Then
                    .ColorBarBottomOn = RGB(255, 0, 0)
                    .ColorBarBottomOff = RGB(200, 0, 0)
                ElseIf InStr(LCase(ctrl.Name), "email") > 0 Then
                    .ColorBarBottomOn = RGB(0, 150, 0)
                    .ColorBarBottomOff = RGB(0, 100, 0)
                End If
            End With
        End If
    Next ctrl
End Sub
```

## Working with Style Collection

### Iterating Through All Style Elements
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Iterate through all style elements and output information
    Dim item As clsModernStyle
    For Each item In MStyleItem.StyleItems
        Debug.Print "Element: " & item.Name & ", Type: " & item.ControlType
    Next item
End Sub
```

### Finding Elements by Criteria
```vba
Private Function FindControlsByType(controlType As String) As Collection
    Dim result As New Collection
    Dim item As clsModernStyle
    
    For Each item In MStyleItem.StyleItems
        If item.ControlType = controlType Then
            result.Add item
        End If
    Next item
    
    Set FindControlsByType = result
End Function

Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Find all text boxes
    Dim textBoxes As Collection
    Set textBoxes = FindControlsByType("TextBox")
    
    ' Apply special style to all text boxes
    Dim item As clsModernStyle
    For Each item In textBoxes
        With item
            .ColorBarBottomOn = RGB(0, 100, 200)
            .ColorBarBottomOff = RGB(150, 150, 150)
        End With
    Next item
End Sub
```

### Group Style Configuration
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' List of control names for group configuration
    Dim importantFields As Variant
    importantFields = Array("UsernameBox", "PasswordBox", "EmailBox")
    
    ' Apply red border to important fields
    Dim i As Integer
    For i = 0 To UBound(importantFields)
        On Error Resume Next
        With MStyleItem.getItemByName(importantFields(i))
            .ColorBarBottomOn = RGB(255, 0, 0)
            .ColorBarBottomOff = RGB(200, 0, 0)
        End With
        On Error GoTo 0
    Next i
End Sub
```

## Integration with Existing Forms

### Adding Styles to Existing Form
```vba
' In a separate module
Public Sub ApplyModernStyleToForm(formName As String)
    Dim frm As Object
    Set frm = VBA.Interaction.CreateObject("Forms." & formName)
    
    Dim style As New clsModernStyle
    style.Initialize frm
    
    frm.Show
End Sub

' Usage
Private Sub CommandButton1_Click()
    ApplyModernStyleToForm "ExistingForm"
End Sub
```

### Gradual Form Styling
```vba
Private Sub UserForm_Initialize()
    ' Initialize without styling
    Set MStyleItem = New clsModernStyle
    
    ' Add controls programmatically
    AddStyledControl "TextBox", "UserInput", 50, 50, 200, 20
    AddStyledControl "ComboBox", "SelectionBox", 50, 80, 200, 20
    AddStyledControl "CheckBox", "AgreementBox", 50, 110, 200, 20
End Sub

Private Sub AddStyledControl(controlType As String, controlName As String, _
                           leftPos As Single, topPos As Single, _
                           widthSize As Single, heightSize As Single)
    Dim newControl As MSForms.control
    
    ' Create control
    Set newControl = Me.Controls.Add("Forms." & controlType & ".1", controlName, True)
    With newControl
        .left = leftPos
        .top = topPos
        .width = widthSize
        .height = heightSize
    End With
    
    ' Initialize styles only for the new control
    ' (requires class modification to support individual styling)
    Call MStyleItem.Initialize(Me)
End Sub
```

## Event Handling Examples

### Handling Events of Styled Elements
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub

Private Sub TextBox1_Change()
    ' Handle text change in styled element
    If Len(TextBox1.Value) > 0 Then
        ' Change style based on content
        With MStyleItem.getItemByName(TextBox1.Name)
            If IsValidEmail(TextBox1.Value) Then
                .ColorBarBottomOn = RGB(0, 150, 0)   ' Green for valid email
            Else
                .ColorBarBottomOn = RGB(255, 0, 0)   ' Red for invalid
            End If
        End With
    End If
End Sub

Private Function IsValidEmail(email As String) As Boolean
    ' Simple email validation
    IsValidEmail = (InStr(email, "@") > 0 And InStr(email, ".") > 0)
End Function
```

### Events for Toggle Switches
```vba
Private Sub ToggleOption1_Change()
    If ToggleOption1.Value Then
        ' Change style of other elements when toggle changes
        With MStyleItem.getItemByName(TextBox2.Name)
            .Enabled = True
            .ColorBackGroundOff = RGB(255, 255, 255)
        End With
    Else
        With MStyleItem.getItemByName(TextBox2.Name)
            .Enabled = False
            .ColorBackGroundOff = RGB(240, 240, 240)
        End With
    End If
End Sub
```

## Creating Themes

### Theme Manager
```vba
' In a separate class or module
Public Type ThemeColors
    TitleActive As Long
    TitleInactive As Long
    BottomLineActive As Long
    BottomLineInactive As Long
    BackgroundActive As Long
    BackgroundInactive As Long
End Type

Public Function GetTheme(themeName As String) As ThemeColors
    Dim theme As ThemeColors
    
    Select Case themeName
        Case "BlueTheme"
            theme.TitleActive = RGB(0, 100, 200)
            theme.TitleInactive = RGB(120, 120, 120)
            theme.BottomLineActive = RGB(0, 100, 200)
            theme.BottomLineInactive = RGB(180, 180, 180)
            theme.BackgroundActive = RGB(255, 255, 255)
            theme.BackgroundInactive = RGB(245, 245, 245)
        Case "GreenTheme"
            theme.TitleActive = RGB(0, 120, 0)
            theme.TitleInactive = RGB(100, 100, 100)
            theme.BottomLineActive = RGB(0, 150, 0)
            theme.BottomLineInactive = RGB(150, 200, 150)
            theme.BackgroundActive = RGB(250, 255, 250)
            theme.BackgroundInactive = RGB(240, 250, 240)
        Case "RedTheme"
            theme.TitleActive = RGB(200, 0, 0)
            theme.TitleInactive = RGB(120, 80, 80)
            theme.BottomLineActive = RGB(200, 0, 0)
            theme.BottomLineInactive = RGB(220, 150, 150)
            theme.BackgroundActive = RGB(255, 250, 250)
            theme.BackgroundInactive = RGB(250, 240, 240)
    End Select
    
    GetTheme = theme
End Function

Private Sub UserForm_Initialize()
    Dim currentTheme As ThemeColors
    currentTheme = GetTheme("BlueTheme")  ' Or "GreenTheme", "RedTheme"
    
    Set MStyleItem = New clsModernStyle
    MStyleItem.Initialize Me, _
        ColorBarTitleOn:=currentTheme.TitleActive, _
        ColorBarTitleOff:=currentTheme.TitleInactive, _
        ColorBarBottomOn:=currentTheme.BottomLineActive, _
        ColorBarBottomOff:=currentTheme.BottomLineInactive, _
        ColorBackGroundOn:=currentTheme.BackgroundActive, _
        ColorBackGroundOff:=currentTheme.BackgroundInactive
End Sub
```

### Theme Switching
```vba
Private Sub ThemeComboBox_Change()
    ApplyTheme ThemeComboBox.Value
End Sub

Private Sub ApplyTheme(themeName As String)
    Dim theme As ThemeColors
    theme = GetTheme(themeName)
    
    ' Apply theme to all elements
    Dim item As clsModernStyle
    For Each item In MStyleItem.StyleItems
        With item
            .ColorBarTitleOn = theme.TitleActive
            .ColorBarTitleOff = theme.TitleInactive
            .ColorBarBottomOn = theme.BottomLineActive
            .ColorBarBottomOff = theme.BottomLineInactive
            .ColorBackGroundOn = theme.BackgroundActive
            .ColorBackGroundOff = theme.BackgroundInactive
        End With
    Next item
End Sub
```

## Advanced Examples

### Login Form with Validation
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Set tooltip texts
    UsernameBox.ControlTipText = "Username"
    PasswordBox.ControlTipText = "Password"
    LoginButton.Caption = "Login"
End Sub

Private Sub UsernameBox_Change()
    ValidateInput "Username", UsernameBox.Value
End Sub

Private Sub PasswordBox_Change()
    ValidateInput "Password", PasswordBox.Value
End Sub

Private Sub ValidateInput(inputType As String, inputValue As String)
    Dim isValid As Boolean
    Dim color As Long
    
    Select Case inputType
        Case "Username"
            isValid = Len(inputValue) >= 3
        Case "Password"
            isValid = Len(inputValue) >= 6
    End Select
    
    color = IIf(isValid, RGB(0, 150, 0), RGB(255, 0, 0))
    
    With MStyleItem.getItemByName(GetControlByName(inputType & "Box").Name)
        .ColorBarBottomOn = color
        .ColorBarBottomOff = IIf(isValid, RGB(200, 200, 200), RGB(255, 200, 200))
    End With
End Sub

Private Function GetControlByName(controlName As String) As MSForms.control
    Set GetControlByName = Me.Controls(controlName)
End Function

Private Sub LoginButton_Click()
    ' Check validity of all fields before submission
    If IsValidForm Then
        MsgBox "Login successful!", vbInformation
        Unload Me
    Else
        MsgBox "Please check the validity of the fields.", vbExclamation
    End If
End Sub

Private Function IsValidForm() As Boolean
    IsValidForm = (Len(UsernameBox.Value) >= 3 And Len(PasswordBox.Value) >= 6)
End Function
```

### Dynamic Form with Field Addition
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Initialize with basic elements
    Call MStyleItem.Initialize(Me)
    
    ' Add dynamic fields
    AddDynamicField "First Name", "firstName", 20, 100
    AddDynamicField "Last Name", "lastName", 20, 140
    AddDynamicField "Email", "email", 20, 180
End Sub

Private Sub AddDynamicField(promptText As String, fieldName As String, _
                          leftPos As Single, topPos As Single)
    ' Create text box
    Dim newTextBox As MSForms.TextBox
    Set newTextBox = Me.Controls.Add("Forms.TextBox.1", fieldName, True)
    
    With newTextBox
        .left = leftPos
        .top = topPos
        .width = 200
        .height = 20
        .ControlTipText = promptText
    End With
    
    ' Reinitialize styles for added elements
    Call MStyleItem.Initialize(Me)
End Sub

Private Sub AddFieldButton_Click()
    ' Add new field on button click
    Static fieldCounter As Integer
    fieldCounter = fieldCounter + 1
    
    AddDynamicField "Additional Field " & fieldCounter, _
                    "extraField" & fieldCounter, 20, 180 + fieldCounter * 40
End Sub
```

These examples demonstrate various ways to use the `clsModernStyle` class to create modern and functional user forms in Excel. Each example can be adapted to specific application requirements.