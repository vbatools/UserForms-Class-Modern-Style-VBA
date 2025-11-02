VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsModernStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class:           clsModernStyle
' Author:          VBATools
' Version:         1.0.8
' Creation Date:   10.10.2025 15:30
' Update Date:     01.11.2025 09:59
'
' Purpose:
'   Class designed to style MSForms controls with a modern design.
'   Implements various visual effects such as focus animation, color and font settings,
'   adding icons and visual elements.
'
' Main Features:
'   - Applying modern style to various controls (TextBox, ComboBox, ListBox, CheckBox, OptionButton, etc.)
'   - Supporting focus animation
'   - Configuring color and font
'   - Adding icons and visual elements
'   - Managing visibility and state of controls
'
' Usage:
'   1. Create an instance of clsModernStyle class
'   2. Call the Initialize method, passing a reference to the UserForm
'   3. The class automatically applies style to all controls on the form
'
' Example Usage:
'   Dim style As New clsModernStyle
'   style.Initialize Me ' where Me is the UserForm
'
' Dependencies:
'   - Microsoft Forms 2.0 Object Library
'   - VBA
'
' License:
'   Apache License
'========================================================================================
'========================================================================================
' Enumeration: enumIcons
' Purpose: Enumeration of icons for various controls
'========================================================================================
Public Enum enumIcons
    ArrowOff = &HE011                       ' Dropdown list arrow (off)
    ArrowOn = &HE010                        ' Dropdown list arrow (on)
    CheckBox1 = 59193                       ' Square (normal)
    Checkbox14 = 61803                      ' Square (small)
    CheckboxComposite = 59194               ' Square with checkmark
    CheckboxComposite14 = 61804             ' Square with checkmark (small)
    CheckboxCompositeReversed = 59197       ' Square with checkmark (reversed)
    CheckboxIndeterminateCombo = 61806      ' Square with dash
    CheckboxIndeterminateCombo14 = 61805    ' Square with dash (small)
    CheckboxFill = 59195                    ' Square (filled)
    CheckMark = 59198                       ' Checkmark
    CircleFill = 59963                      ' Circle (filled)
    CircleRing = 59962                      ' Circle (outline)
    FavoriteStar = 59188                    ' Star (normal)
    FavoriteStarFill = 59189                ' Star (filled)
    Heart = 60241                           ' Heart (normal)
    HeartFill = 60242                       ' Heart (filled)
    InkingColorFill = 6075                 ' Brush (filled)
    InkingColorOutline = 60774              ' Brush (outline)
    PaginationDotOutline10 = 61734          ' Dot (outline)
    PaginationDotSolid10 = 61735            ' Dot (filled)
    PasswordChar = 149                      ' Character for hiding password
    RadioBtnOff = 60618                     ' Radio button (off)
    RadioBtnOn = 60619                      ' Radio button (on)
    ToggleOff = 60434                       ' Toggle switch (off)
    ToggleOn = 60433                        ' Toggle switch (on)
    ToggleThumb = 60436                     ' Toggle switch thumb
End Enum
'================================================================================
' Section: Global variables and constants
'================================================================================
' Collection of all style items
Private mItems      As Collection
' Font constants
Private Const FONT_NAME_ICON As String = "Segoe MDL2 Assets"
' Constants for control types
Private Const CONTROL_TYPE_TEXTBOX As String = "TextBox"
Private Const CONTROL_TYPE_COMBOBOX As String = "ComboBox"
Private Const CONTROL_TYPE_LISTBOX As String = "ListBox"
Private Const CONTROL_TYPE_CHECKBOX As String = "CheckBox"
Private Const CONTROL_TYPE_OPTIONBUTTON As String = "OptionButton"
Private Const CONTROL_TYPE_FRAME As String = "Frame"
Private Const CONTROL_TYPE_LABEL As String = "Label"
Private Const CONTROL_TYPE_COMMANDBUTTON As String = "CommandButton"
Private Const CONTROL_TYPE_MULTI_PAGE As String = "MultiPage"
Private Const CONTROL_TYPE_IMAGE As String = "Image"
Private Const CONTROL_TYPE_TABSTRIP As String = "TabStrip"
Private Const CONTROL_TYPE_SCROLLBAR As String = "ScrollBar"
Private Const CONTROL_TYPE_SPINBUTTON As String = "SpinButton"
' Constants for additional control names
Private Const BAR_BOTTOM As String = "_barBottom"
Private Const BAR_TITLE As String = "_barTitle"
Private Const BAR_ICON As String = "_barIcon"
Private Const BACK_GROUND As String = "_BackGround"
Private Const DROP_ARROW As String = "_DropArrow"
Private Const BTN_CLEAR As String = "_BtnClear"
' Constants for control behavior
Private Const CONTROL_SWITCH As String = "SWITCH"
' Objects with event handlers
Private WithEvents myForm As MSForms.UserForm
Attribute myForm.VB_VarHelpID = -1
Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mComboBox As MSForms.ComboBox
Attribute mComboBox.VB_VarHelpID = -1
Private WithEvents mListBox As MSForms.ListBox
Attribute mListBox.VB_VarHelpID = -1
Private WithEvents mFrame As MSForms.Frame
Attribute mFrame.VB_VarHelpID = -1
'Private WithEvents mPage As MSForms.MultiPage
Private WithEvents mLabel As MSForms.label
Attribute mLabel.VB_VarHelpID = -1
Private WithEvents mCommandButton As MSForms.CommandButton
Attribute mCommandButton.VB_VarHelpID = -1
Private WithEvents mCheckBox As MSForms.CheckBox
Attribute mCheckBox.VB_VarHelpID = -1
Private WithEvents mOptionButton As MSForms.OptionButton
Attribute mOptionButton.VB_VarHelpID = -1
'================================================================================
' Section: Additional controls for styling
'================================================================================
Private mBarBottom  As MSForms.label
Private mBarTitle   As MSForms.label
Private mBarIcon    As MSForms.label
Private mBackGround As MSForms.label
Private WithEvents mDropArrow As MSForms.label
Attribute mDropArrow.VB_VarHelpID = -1
Private WithEvents mBtnClear As MSForms.label
Attribute mBtnClear.VB_VarHelpID = -1
Private WithEvents mTgBorder As MSForms.label
Attribute mTgBorder.VB_VarHelpID = -1
Private WithEvents mChkBoxBtn As MSForms.label
Attribute mChkBoxBtn.VB_VarHelpID = -1
Private WithEvents mChkBoxCaption As MSForms.label
Attribute mChkBoxCaption.VB_VarHelpID = -1
'================================================================================
' Section: Main control properties
'================================================================================
' Main control variables
Private mControl    As MSForms.control
Private mType       As String
Private mName       As String
Private mControlTipText As String
Private mVisible    As Boolean
Private mLocked     As Boolean
Private mEnabled    As Boolean
' Positioning and sizes
Private mTop        As Single
Private mLeft       As Single
Private mHeight     As Single
Private mWidth      As Single
'================================================================================
' Section: Font properties
'================================================================================
Private mFontSizeTitleOff As Integer
Private mFontSizeTitleOn As Integer
Private mFontName   As String
'================================================================================
' Section: Color properties for various elements
'================================================================================
' Colors for various elements
Private mColorBarTitleOn As XlRgbColor
Private mColorBarTitleOff As XlRgbColor
Private mColorBarBottomOn As XlRgbColor
Private mColorBarBottomOff As XlRgbColor
Private mColorBackGroundOn As XlRgbColor
Private mColorBackGroundOff As XlRgbColor
Private mColorBarIconOn As XlRgbColor
Private mColorBarIconOff As XlRgbColor
Private mColorDropArrowOn As XlRgbColor
Private mColorDropArrowOff As XlRgbColor
' Colors for toggles
Private mColorTgBorderOn As XlRgbColor
Private mColorTgBorderOff As XlRgbColor
Private mColorChkBoxBtnOn As XlRgbColor
Private mColorChkBoxBtnOff As XlRgbColor
Private mColorChkBoxCaptionOn As XlRgbColor
Private mColorChkBoxCaptionOff As XlRgbColor
'================================================================================
' Section: Character properties for arrows and switches
'================================================================================
Private mChrDropArrowOn As String
Private mChrDropArrowOff As String
Private mChrChkBoxBtnOn As String
Private mChrChkBoxBtnOff As String
Private mChrOptBoxBtnOn As String
Private mChrOptBoxBtnOff As String
'================================================================================
' Section: Properties for item collection
'================================================================================
Public Property Get Items() As Collection
    Set Items = mItems
End Property

Public Property Set Items(ByRef Items As Collection)
    Set mItems = Items
End Property
'================================================================================
' Section: Properties for main control
'================================================================================
Public Property Get control() As MSForms.control
    Set control = mControl
End Property

Public Property Set control(ByRef control As MSForms.control)
    Set mControl = control
End Property
'================================================================================
' Section: Properties for various control types
'================================================================================
Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = mTextBox
End Property

Public Property Set TextBox(ByRef TextBox As MSForms.TextBox)
    Set mTextBox = TextBox
End Property

Public Property Get ComboBox() As MSForms.ComboBox
    Set ComboBox = mComboBox
End Property

Public Property Set ComboBox(ByRef ComboBox As MSForms.ComboBox)
    Set mComboBox = ComboBox
End Property

Public Property Get ListBox() As MSForms.ListBox
    Set ListBox = mListBox
End Property

Public Property Set ListBox(ByRef ListBox As MSForms.ListBox)
    Set mListBox = ListBox
End Property

Public Property Get CheckBox() As MSForms.CheckBox
    Set CheckBox = mCheckBox
End Property

Public Property Set CheckBox(ByRef CheckBox As MSForms.CheckBox)
    Set mCheckBox = CheckBox
End Property

Public Property Get OptionButton() As MSForms.OptionButton
    Set OptionButton = mOptionButton
End Property

Public Property Set OptionButton(ByRef OptionButton As MSForms.OptionButton)
    Set mOptionButton = OptionButton
End Property

Public Property Get Frame() As MSForms.Frame
    Set Frame = mFrame
End Property

Public Property Set Frame(ByRef Frame As MSForms.Frame)
    Set mFrame = Frame
End Property

Public Property Get label() As MSForms.label
    Set label = mLabel
End Property

Public Property Set label(ByRef label As MSForms.label)
    Set mLabel = label
End Property

Public Property Get CommandButton() As MSForms.CommandButton
    Set CommandButton = mCommandButton
End Property

Public Property Set CommandButton(ByRef CommandButton As MSForms.CommandButton)
    Set mCommandButton = CommandButton
End Property
'================================================================================
' Section: Properties for control type and name
'================================================================================
Public Property Get ControlType() As String
    ControlType = mType
End Property

Public Property Let ControlType(ByVal NameType As String)
    mType = NameType
End Property

Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal Name As String)
    mName = Name
End Property
'================================================================================
' Section: Properties for visibility, locking and availability
'================================================================================
' Properties for visibility
Public Property Get Visible() As Boolean
    Visible = mVisible
End Property
Public Property Let Visible(ByRef Visible As Boolean)
    mVisible = Visible
    If Not mControl Is Nothing Then mControl.Visible = Visible
    Call SetControlVisibility(mBarBottom, Visible)
    Call SetControlVisibility(mBarTitle, Visible)
    Call SetControlVisibility(mBarIcon, Visible)
    Call SetControlVisibility(mBackGround, Visible)
    Call SetControlVisibility(mTgBorder, Visible)
    Call SetControlVisibility(mChkBoxBtn, Visible)
    Call SetControlVisibility(mChkBoxCaption, Visible)
    Call SetControlVisibility(mDropArrow, Visible)
    Call SetControlVisibility(mBtnClear, Visible)
End Property
' Properties for locking state
Public Property Get Locked() As Boolean
    Locked = mLocked
End Property

Public Property Let Locked(ByRef Locked As Boolean)
    mLocked = Locked
    If Not mControl Is Nothing Then mControl.Locked = Locked
    Call SetControlLock(mBarBottom, Locked)
    Call SetControlLock(mBarTitle, Locked)
    Call SetControlLock(mBarIcon, Locked)
    Call SetControlLock(mBackGround, Locked)
    Call SetControlLock(mTgBorder, Locked)
    Call SetControlLock(mChkBoxBtn, Locked)
    Call SetControlLock(mChkBoxCaption, Locked)
    Call SetControlLock(mDropArrow, Locked)
    Call SetControlLock(mBtnClear, Locked)
End Property
' Properties for availability state
Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByRef Enabled As Boolean)
    mEnabled = Enabled
    If Not mControl Is Nothing Then mControl.Enabled = Enabled
    Call SetControlEnabled(mBarBottom, Enabled)
    Call SetControlEnabled(mBarTitle, Enabled)
    Call SetControlEnabled(mBarIcon, Enabled)
    Call SetControlEnabled(mBackGround, Enabled)
    Call SetControlEnabled(mTgBorder, Enabled)
    Call SetControlEnabled(mChkBoxBtn, Enabled)
    Call SetControlEnabled(mChkBoxCaption, Enabled)
    Call SetControlEnabled(mDropArrow, Enabled)
    Call SetControlEnabled(mBtnClear, Enabled)
End Property
'---------------------------------------------------------------------------------------
' Procedure: SetControlEnabled
' Purpose: Internal method to set the availability state of a control
' Parameters:
'   control - Control to change availability state
'   isEnabled - Flag for availability state
'---------------------------------------------------------------------------------------
Private Sub SetControlEnabled(ByRef control As MSForms.control, ByVal isEnabled As Boolean)
    If Not control Is Nothing Then control.Enabled = isEnabled
End Sub
'---------------------------------------------------------------------------------------
' Procedure: SetControlVisibility
' Purpose: Internal method to set the visibility of a control
' Parameters:
'   control - Control to change visibility
'   isVisible - Visibility flag
'---------------------------------------------------------------------------------------
Private Sub SetControlVisibility(ByRef control As MSForms.control, ByVal isVisible As Boolean)
    If Not control Is Nothing Then control.Visible = isVisible
End Sub
'---------------------------------------------------------------------------------------
' Procedure: SetControlLock
' Purpose: Internal method to set the locking state of a control
' Parameters:
'   control - Control to change locking state
'   isLocked - Flag for locking state
'---------------------------------------------------------------------------------------
Private Sub SetControlLock(ByRef control As MSForms.control, ByVal isLocked As Boolean)
    If Not control Is Nothing Then control.Locked = isLocked
End Sub
'================================================================================
' Section: Properties for tooltip
'================================================================================
Public Property Get ControlTipText() As String
    ControlTipText = mControlTipText
End Property

Public Property Let ControlTipText(ByVal ControlTipText As String)
    mControlTipText = ControlTipText
End Property
'================================================================================
' Section: Properties for style element colors
'================================================================================
Public Property Get ColorBarTitleOn() As XlRgbColor
    ColorBarTitleOn = mColorBarTitleOn
End Property

Public Property Let ColorBarTitleOn(ByVal Color As XlRgbColor)
    mColorBarTitleOn = Color
End Property

Public Property Get ColorBarBottomOn() As XlRgbColor
    ColorBarBottomOn = mColorBarBottomOn
End Property

Public Property Let ColorBarBottomOn(ByVal Color As XlRgbColor)
    mColorBarBottomOn = Color
End Property

Public Property Get ColorBarBottomOff() As XlRgbColor
    ColorBarBottomOff = mColorBarBottomOff
End Property

Public Property Let ColorBarBottomOff(ByVal Color As XlRgbColor)
    mColorBarBottomOff = Color
End Property

Public Property Get ColorBarTitleOff() As XlRgbColor
    ColorBarTitleOff = mColorBarTitleOff
End Property

Public Property Let ColorBarTitleOff(ByVal Color As XlRgbColor)
    mColorBarTitleOff = Color
End Property

Public Property Get ColorBackGroundOn() As XlRgbColor
    ColorBackGroundOn = mColorBackGroundOn
End Property

Public Property Let ColorBackGroundOn(ByVal Color As XlRgbColor)
    mColorBackGroundOn = Color
End Property

Public Property Get ColorBackGroundOff() As XlRgbColor
    ColorBackGroundOff = mColorBackGroundOff
End Property

Public Property Let ColorBackGroundOff(ByVal Color As XlRgbColor)
    mColorBackGroundOff = Color
End Property

Public Property Get ColorBarIconOn() As XlRgbColor
    ColorBarIconOn = mColorBarIconOn
End Property

Public Property Let ColorBarIconOn(ByVal Color As XlRgbColor)
    mColorBarIconOn = Color
End Property

Public Property Get ColorBarIconOff() As XlRgbColor
    ColorBarIconOff = mColorBarIconOff
End Property

Public Property Let ColorBarIconOff(ByVal Color As XlRgbColor)
    mColorBarIconOff = Color
End Property

Public Property Get ColorDropArrowOn() As XlRgbColor
    ColorDropArrowOn = mColorDropArrowOn
End Property

Public Property Let ColorDropArrowOn(ByVal Color As XlRgbColor)
    mColorDropArrowOn = Color
End Property

Public Property Get ColorDropArrowOff() As XlRgbColor
    ColorDropArrowOff = mColorDropArrowOff
End Property

Public Property Let ColorDropArrowOff(ByVal Color As XlRgbColor)
    mColorDropArrowOff = Color
End Property
'================================================================================
' Section: Properties for switch colors
'================================================================================
Public Property Let ColorTgBorderOn(ByVal Color As XlRgbColor)
    mColorTgBorderOn = Color
End Property

Public Property Get ColorTgBorderOn() As XlRgbColor
    ColorTgBorderOn = mColorTgBorderOn
End Property

Public Property Let ColorTgBorderOff(ByVal Color As XlRgbColor)
    mColorTgBorderOff = Color
End Property

Public Property Get ColorTgBorderOff() As XlRgbColor
    ColorTgBorderOff = mColorTgBorderOff
End Property

Public Property Let ColorChkBoxBtnOn(ByVal Color As XlRgbColor)
    mColorChkBoxBtnOn = Color
End Property

Public Property Get ColorChkBoxBtnOn() As XlRgbColor
    ColorChkBoxBtnOn = mColorChkBoxBtnOn
End Property

Public Property Let ColorChkBoxBtnOff(ByVal Color As XlRgbColor)
    mColorChkBoxBtnOff = Color
End Property

Public Property Get ColorChkBoxBtnOff() As XlRgbColor
    ColorChkBoxBtnOff = mColorChkBoxBtnOff
End Property

Public Property Let ColorTgFore(ByVal Color As XlRgbColor)
    If Not mChkBoxCaption Is Nothing Then mChkBoxCaption.ForeColor = Color
End Property

Public Property Get ColorTgFore() As XlRgbColor
    If Not mChkBoxCaption Is Nothing Then ColorTgFore = mChkBoxCaption.ForeColor
End Property

Public Property Let ColorChkBoxCaptionOn(ByVal Color As Long)
    mColorChkBoxCaptionOn = Color
End Property

Public Property Get ColorChkBoxCaptionOn() As Long
    ColorChkBoxCaptionOn = mColorChkBoxCaptionOn
End Property

Public Property Let ColorChkBoxCaptionOff(ByVal Color As Long)
    mColorChkBoxCaptionOff = Color
End Property

Public Property Get ColorChkBoxCaptionOff() As Long
    ColorChkBoxCaptionOff = mColorChkBoxCaptionOff
End Property
'================================================================================
' Section: Properties for arrow characters
'================================================================================
Public Property Let ChrDropArrowOn(ByVal chr As String)
    mChrDropArrowOn = chr
End Property

Public Property Get ChrDropArrowOn() As String
    ChrDropArrowOn = mChrDropArrowOn
End Property

Public Property Let ChrDropArrowOff(ByVal chr As String)
    mChrDropArrowOff = chr
End Property

Public Property Get ChrDropArrowOff() As String
    ChrDropArrowOff = mChrDropArrowOff
End Property

Public Property Let ChrChkBoxBtnOn(ByVal Code As String)
    mChrChkBoxBtnOn = Code
End Property

Public Property Get ChrChkBoxBtnOn() As String
    ChrChkBoxBtnOn = mChrChkBoxBtnOn
End Property

Public Property Let ChrChkBoxBtnOff(ByVal Code As String)
    mChrChkBoxBtnOff = Code
End Property

Public Property Get ChrChkBoxBtnOff() As String
    ChrChkBoxBtnOff = mChrChkBoxBtnOff
End Property

Public Property Let ChrOptBoxBtnOn(ByVal Code As String)
    mChrOptBoxBtnOn = Code
End Property

Public Property Get ChrOptBoxBtnOn() As String
    ChrOptBoxBtnOn = mChrOptBoxBtnOn
End Property

Public Property Let ChrOptBoxBtnOff(ByVal Code As String)
    mChrOptBoxBtnOff = Code
End Property

Public Property Get ChrOptBoxBtnOff() As String
    ChrOptBoxBtnOff = mChrOptBoxBtnOff
End Property
'================================================================================
' Section: Properties for positioning and sizes
'================================================================================
Public Property Get top() As Single
    top = mTop
End Property

Public Property Let top(ByVal Size As Single)
    mTop = Size
End Property

Public Property Get left() As Single
    left = mLeft
End Property

Public Property Let left(ByVal Size As Single)
    mLeft = Size
End Property

Public Property Get height() As Single
    height = mHeight
End Property

Public Property Let height(ByVal Size As Single)
    mHeight = Size
End Property

Public Property Get width() As Single
    width = mWidth
End Property

Public Property Let width(ByVal Size As Single)
    mWidth = Size
End Property
'================================================================================
' Section: Properties for font
'================================================================================
Public Property Get FontSizeTitleOff() As Integer
    FontSizeTitleOff = mFontSizeTitleOff
End Property

Public Property Let FontSizeTitleOff(ByVal Size As Integer)
    mFontSizeTitleOff = Size
End Property

Public Property Get FontSizeTitleOn() As Integer
    FontSizeTitleOn = mFontSizeTitleOn
End Property

Public Property Let FontSizeTitleOn(ByVal Size As Integer)
    mFontSizeTitleOn = Size
End Property

Public Property Get FontName() As String
    FontName = mFontName
End Property

Public Property Let FontName(ByVal Name As String)
    mFontName = Name
End Property
'================================================================================
' Section: Properties for additional style elements
'================================================================================
Public Property Get BarBottom() As MSForms.label
    Set BarBottom = mBarBottom
End Property

Public Property Set BarBottom(ByRef label As MSForms.label)
    Set mBarBottom = label
End Property

Public Property Set BarTitle(ByRef label As MSForms.label)
    Set mBarTitle = label
End Property

Public Property Get BarTitle() As MSForms.label
    Set BarTitle = mBarTitle
End Property

Public Property Set BarIcon(ByRef label As MSForms.label)
    Set mBarIcon = label
End Property

Public Property Get BarIcon() As MSForms.label
    Set BarIcon = mBarIcon
End Property

Public Property Set BackGround(ByRef label As MSForms.label)
    Set mBackGround = label
End Property

Public Property Get BackGround() As MSForms.label
    Set BackGround = mBackGround
End Property

Public Property Set DropArrow(ByRef label As MSForms.label)
    Set mDropArrow = label
End Property

Public Property Get DropArrow() As MSForms.label
    Set DropArrow = mDropArrow
End Property

Public Property Set BtnClear(ByRef label As MSForms.label)
    Set mBtnClear = label
End Property

Public Property Get BtnClear() As MSForms.label
    Set BtnClear = mBtnClear
End Property

Public Property Set TgBorder(ByRef label As MSForms.label)
    Set mTgBorder = label
End Property

Public Property Get TgBorder() As MSForms.label
    Set TgBorder = mTgBorder
End Property

Public Property Set ChkBoxBtn(ByRef label As MSForms.label)
    Set mChkBoxBtn = label
End Property

Public Property Get ChkBoxBtn() As MSForms.label
    Set ChkBoxBtn = mChkBoxBtn
End Property

Public Property Set ChkBoxCaption(ByRef label As MSForms.label)
    Set mChkBoxCaption = label
End Property

Public Property Get ChkBoxCaption() As MSForms.label
    Set ChkBoxCaption = mChkBoxCaption
End Property

Public Property Get Count() As Byte
    Count = mItems.Count
End Property

Public Property Get getItemByIndex(ByVal index As Integer) As clsModernStyle
    On Error GoTo endGetItem
    Set getItemByIndex = mItems(index)
    Exit Property
endGetItem:
    Err.Clear
End Property

Public Property Get getItemByName(ByVal Name As String) As clsModernStyle
    On Error GoTo endGetItem
    Set getItemByName = mItems(Name)
    Exit Property
endGetItem:
    Err.Clear
End Property

Public Property Get Version() As String
    Version = "Version: 1.0.6" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Date of creation: 10.10.2025 15:30" & vbNewLine & _
            "Date of update: 01.11.2025 09:59"
End Property

'================================================================================
' Section: Initialization and management methods
'================================================================================
'---------------------------------------------------------------------------------------
' Procedure: Initialize
' Purpose: Initialize style for all form controls
' Parameters:
'   Form - Reference to UserForm to which style is applied
'   ColorBarTitleOn - Title color in active state
'   ColorBarTitleOff - Title color in inactive state
'   ColorBarBottomOn - Bottom line color in active state
'   ColorBarBottomOff - Bottom line color in inactive state
'   ColorBackGroundOn - Background color in active state
'   ColorBackGroundOff - Background color in inactive state
'   ColorBarIconOn - Icon color in active state
'   ColorBarIconOff - Icon color in inactive state
'   ColorDropArrowOn - Dropdown arrow color in active state
'   ColorDropArrowOff - Dropdown arrow color in inactive state
'   ColorTgBorderOn - Switch border color in active state
'   ColorTgBorderOff - Switch border color in inactive state
'   ColorChkBoxBtnOn - Checkbox button color in active state
'   ColorChkBoxBtnOff - Checkbox button color in inactive state
'   ChrDropArrowOn - Dropdown arrow character in active state
'   ChrDropArrowOff - Dropdown arrow character in inactive state
'   ColorChkBoxCaptionOn - Checkbox caption color in active state
'   ColorChkBoxCaptionOff - Checkbox caption color in inactive state
'   ChrChkBoxBtnOn - Checkbox button character in active state
'   ChrChkBoxBtnOff - Checkbox button character in inactive state
'---------------------------------------------------------------------------------------
Public Sub Initialize(ByRef Form As MSForms.UserForm, _
        Optional ColorBarTitleOn As XlRgbColor = 14854934, _
        Optional ColorBarTitleOff As XlRgbColor = 10395294, _
        Optional ColorBarBottomOn As XlRgbColor = 14854934, _
        Optional ColorBarBottomOff As XlRgbColor = 10395294, _
        Optional ColorBackGroundOn As XlRgbColor = vbWhite, _
        Optional ColorBackGroundOff As XlRgbColor = 16447476, _
        Optional ColorBarIconOn As XlRgbColor = 14854934, _
        Optional ColorBarIconOff As XlRgbColor = 10395294, _
        Optional ColorDropArrowOn As XlRgbColor = vbBlack, _
        Optional ColorDropArrowOff As XlRgbColor = 10395294, _
        Optional ColorTgBorderOn As XlRgbColor = 14854934, _
        Optional ColorTgBorderOff As XlRgbColor = 10395294, _
        Optional ColorChkBoxBtnOn As XlRgbColor = vbBlack, _
        Optional ColorChkBoxBtnOff As XlRgbColor = 10395294, _
        Optional ChrDropArrowOn As enumIcons = ArrowOn, _
        Optional ChrDropArrowOff As enumIcons = ArrowOff, _
        Optional ColorChkBoxCaptionOn As XlRgbColor = 14854934, _
        Optional ColorChkBoxCaptionOff As XlRgbColor = 10395294, _
        Optional ChrChkBoxBtnOn As enumIcons = CheckboxComposite, _
        Optional ChrChkBoxBtnOff As enumIcons = CheckBox1, _
        Optional ChrOptBoxBtnOn As enumIcons = CircleFill, _
        Optional ChrOptBoxBtnOff As enumIcons = CircleRing)

    On Error GoTo ErrorHandler
    Set myForm = Form
    Set mItems = New Collection
    Dim cntrl       As MSForms.control
    Dim itemStyle   As clsModernStyle

    ' Loop through all controls on the form
    For Each cntrl In myForm.Controls
        Set itemStyle = New clsModernStyle
        With itemStyle
            Set .control = cntrl
            .ControlType = TypeName(.control)

            .Visible = .control.Visible
            .Enabled = .control.Enabled
            Select Case .ControlType
                Case CONTROL_TYPE_LABEL, CONTROL_TYPE_FRAME, CONTROL_TYPE_MULTI_PAGE, _
                        CONTROL_TYPE_IMAGE, CONTROL_TYPE_TABSTRIP, CONTROL_TYPE_SCROLLBAR, CONTROL_TYPE_SPINBUTTON
                Case Else
                    .Locked = .control.Locked
            End Select

            ' Setting default colors
            .ColorBarTitleOn = ColorBarTitleOn
            .ColorBarTitleOff = ColorBarTitleOff
            .ColorBarBottomOn = ColorBarBottomOn
            .ColorBarBottomOff = ColorBarBottomOff
            .ColorBackGroundOn = ColorBackGroundOn
            .ColorBackGroundOff = ColorBackGroundOff
            .ColorBarIconOn = ColorBarIconOn
            .ColorBarIconOff = ColorBarIconOff
            .ColorDropArrowOn = ColorDropArrowOn
            .ColorDropArrowOff = ColorDropArrowOff

            .ColorTgBorderOff = ColorTgBorderOff
            .ColorTgBorderOn = ColorTgBorderOn
            .ColorChkBoxBtnOff = ColorChkBoxBtnOff
            .ColorChkBoxBtnOn = ColorChkBoxBtnOn

            .ColorChkBoxCaptionOn = ColorChkBoxCaptionOn
            .ColorChkBoxCaptionOff = ColorChkBoxCaptionOff

            ' Setting characters for arrow
            .ChrDropArrowOn = VBA.ChrW(ChrDropArrowOn)
            .ChrDropArrowOff = VBA.ChrW$(ChrDropArrowOff)

            .ChrChkBoxBtnOff = VBA.ChrW(ChrChkBoxBtnOff)
            .ChrChkBoxBtnOn = VBA.ChrW(ChrChkBoxBtnOn)

            .ChrOptBoxBtnOff = VBA.ChrW(ChrOptBoxBtnOff)
            .ChrOptBoxBtnOn = VBA.ChrW(ChrOptBoxBtnOn)

            ' Setting main properties
            .ControlTipText = .control.ControlTipText
            .Name = .control.Name
            .top = .control.top
            .left = .control.left
            .height = .control.height
            .width = .control.width

            Select Case .ControlType
                Case CONTROL_TYPE_IMAGE, CONTROL_TYPE_SCROLLBAR, CONTROL_TYPE_SPINBUTTON
                Case Else:
                    .FontSizeTitleOff = .control.Font.Size
                    .FontSizeTitleOn = .FontSizeTitleOff * 0.95
                    .FontName = .control.Font.Name
            End Select

            ' Applying style depending on control type
            Call ApplyControlStyle(itemStyle)

            Call mItems.Add(itemStyle, .Name)
            Set .Items = mItems

            ' Activating style for elements with already entered values
            Select Case .ControlType
                Case CONTROL_TYPE_TEXTBOX, CONTROL_TYPE_COMBOBOX, CONTROL_TYPE_LISTBOX
                    If .control.Value <> vbNullString Then Call HandleEnterEvent
            End Select
        End With
    Next cntrl
    Exit Sub

ErrorHandler:
    ' Error handling
    MsgBox "Error initializing style: " & Err.Description, vbCritical
End Sub
'================================================================================
' Section: Methods for setting control styles
'================================================================================
'---------------------------------------------------------------------------------------
' Procedure: ApplyControlStyle
' Purpose: Method to apply style depending on control type
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub ApplyControlStyle(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Select Case .ControlType
            Case CONTROL_TYPE_TEXTBOX
                Set .TextBox = itemStyle.control
                Call setTextBoxStyle(itemStyle)
                Call addBarBottom(itemStyle)
                Call addBarTitle(itemStyle)
                Call addBarIcon(itemStyle)
                Call addBackGround(itemStyle)
                Call addBtnClear(itemStyle)
            Case CONTROL_TYPE_COMBOBOX
                Set .ComboBox = itemStyle.control
                Call setComboBoxStyle(itemStyle)
                Call addBarBottom(itemStyle)
                Call addBarTitle(itemStyle)
                Call addBarIcon(itemStyle)
                Call addBackGround(itemStyle)
                Call addDropArrow(itemStyle)
                Call addBtnClear(itemStyle)
            Case CONTROL_TYPE_LISTBOX
                Set .ListBox = itemStyle.control
                Call setListBoxStyle(itemStyle)
                Call addBarBottom(itemStyle)
                Call addBarTitle(itemStyle)
                Call addBarIcon(itemStyle)
            Case CONTROL_TYPE_CHECKBOX
                Set .CheckBox = itemStyle.control
                If .control.Tag = CONTROL_SWITCH Then
                    Call addCheckBoxSwitch(itemStyle)
                Else
                    Call addCheckBox(itemStyle, .ChrChkBoxBtnOff, .ChrChkBoxBtnOn)
                End If
            Case CONTROL_TYPE_OPTIONBUTTON
                Set .OptionButton = itemStyle.control
                If .control.Tag = CONTROL_SWITCH Then
                    Call addCheckBoxSwitch(itemStyle)
                Else
                    Call addCheckBox(itemStyle, .ChrOptBoxBtnOff, .ChrOptBoxBtnOn)
                End If
            Case CONTROL_TYPE_FRAME
                Set .Frame = itemStyle.control
            Case CONTROL_TYPE_LABEL
                Set .label = itemStyle.control
            Case CONTROL_TYPE_COMMANDBUTTON
                Set .CommandButton = itemStyle.control
        End Select
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: SetCommonStyleProperties
' Purpose: Setting common style properties for a control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub SetCommonStyleProperties(ByRef itemStyle As clsModernStyle)
    With itemStyle
        With .control
            .BackStyle = fmBackStyleTransparent
            .BorderStyle = fmBorderStyleNone
            .SpecialEffect = fmSpecialEffectFlat
            .SelectionMargin = False
            .Font.Name = itemStyle.FontName
            .Font.Size = itemStyle.FontSizeTitleOff
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: setTextBoxStyle
' Purpose: Setting style for text box
' Parameters:
'   itemStyle - Style object for configuration
'   sPasswordChar - Character for displaying password field (optional parameter)
'---------------------------------------------------------------------------------------
Private Sub setTextBoxStyle(ByRef itemStyle As clsModernStyle, _
        Optional sPasswordChar As String = vbNullString)
    Call SetCommonStyleProperties(itemStyle)
    With itemStyle
        With .control
            If VBA.LCase$(itemStyle.Name) Like "*password*" Then
                If sPasswordChar = vbNullString Then sPasswordChar = VBA.chr$(sPasswordChar)
                .PasswordChar = sPasswordChar
            End If
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: setComboBoxStyle
' Purpose: Setting style for combo box
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub setComboBoxStyle(ByRef itemStyle As clsModernStyle)
    Call SetCommonStyleProperties(itemStyle)
    With itemStyle
        With .control
            .BackColor = itemStyle.ColorBackGroundOff
            .ShowDropButtonWhen = fmShowDropButtonWhenNever
            .height = 20
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: setListBoxStyle
' Purpose: Setting style for list box
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub setListBoxStyle(ByRef itemStyle As clsModernStyle)
    With itemStyle
        With .control
            .BorderStyle = fmBorderStyleSingle
            .IntegralHeight = False
            .Font.Name = itemStyle.FontName
            .BackColor = itemStyle.ColorBackGroundOff
            .Font.Size = itemStyle.FontSizeTitleOff
        End With
    End With
End Sub
'Adding control styles * * * * * *
'================================================================================
' Section: Methods for adding styles to controls
'================================================================================
'---------------------------------------------------------------------------------------
' Function: CreateStyledLabel
' Purpose: Creating and configuring main properties of additional element
' Parameters:
'   itemStyle - Style object for configuration
'   controlName - Name of control
'   zIndex - Element placement order (default 1)
' Returns: New MSForms.label control
'---------------------------------------------------------------------------------------
Private Function CreateStyledLabel(ByRef itemStyle As clsModernStyle, _
        ByVal controlName As String, _
        Optional ByVal zIndex As Integer = 1) As MSForms.label
    Set CreateStyledLabel = itemStyle.control.Parent.Controls.Add("Forms.Label.1", itemStyle.Name & controlName)
    With CreateStyledLabel
        .Visible = itemStyle.Visible
        .ZOrder zIndex
    End With
End Function
'---------------------------------------------------------------------------------------
' Procedure: SetCommonFontProperties
' Purpose: Setting common font properties for a control
' Parameters:
'   label - MSForms.label control
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub SetCommonFontProperties(ByRef label As MSForms.label, ByRef itemStyle As clsModernStyle)
    With label
        .Font.Name = itemStyle.FontName
        .Font.Size = itemStyle.FontSizeTitleOff
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addBarBottom
' Purpose: Adding bottom style line for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addBarBottom(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .BarBottom = CreateStyledLabel(itemStyle, BAR_BOTTOM, 1)
        Call ConfigureStyleElement(.BarBottom, _
                itemStyle.width, _
                1, _
                itemStyle.left, _
                itemStyle.top + itemStyle.height)
        With .BarBottom
            .BackColor = itemStyle.ColorBarBottomOff
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addBarTitle
' Purpose: Adding style title for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addBarTitle(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .BarTitle = CreateStyledLabel(itemStyle, BAR_TITLE, 1)
        Dim titleLeft As Single, titleTop As Single
        titleLeft = itemStyle.left + 2
        titleTop = itemStyle.top
        If itemStyle.ControlType = CONTROL_TYPE_LISTBOX Then
            titleTop = itemStyle.top - 17
            titleLeft = itemStyle.left
        End If
        Call ConfigureStyleElement(itemStyle.BarTitle, _
                itemStyle.width, _
                itemStyle.height, _
                titleLeft, _
                titleTop)
        With .BarTitle
            .Caption = itemStyle.ControlTipText
            .ForeColor = itemStyle.ColorBarTitleOff
            Call SetCommonFontProperties(itemStyle.BarTitle, itemStyle)
            .BackStyle = fmBackStyleTransparent
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addBarIcon
' Purpose: Adding style icon for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addBarIcon(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .BarIcon = CreateStyledLabel(itemStyle, BAR_ICON, 1)
        Call ConfigureStyleElement(itemStyle.BarIcon, _
                20, _
                20, _
                itemStyle.left - 20, _
                itemStyle.top + 2)
        With .BarIcon
            .Font.Name = FONT_NAME_ICON    '"icon-font" '
            .Font.Size = 15
            .Font.Bold = True
            .TextAlign = fmTextAlignCenter
            If itemStyle.control.Tag <> vbNullString Then .Caption = VBA.ChrW$(itemStyle.control.Tag)
            .ForeColor = itemStyle.ColorBarIconOff
            .BackStyle = fmBackStyleTransparent
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addBackGround
' Purpose: Adding style background for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addBackGround(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .BackGround = CreateStyledLabel(itemStyle, BACK_GROUND, 1)
        Call ConfigureStyleElement(itemStyle.BackGround, _
                itemStyle.width, _
                itemStyle.height + 4, _
                itemStyle.left, _
                itemStyle.top - 4)
        With .BackGround
            .BackColor = itemStyle.ColorBackGroundOff
            Call SetCommonFontProperties(itemStyle.BackGround, itemStyle)
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addDropArrow
' Purpose: Adding dropdown arrow style for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addDropArrow(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .DropArrow = CreateStyledLabel(itemStyle, DROP_ARROW, 0)
        Call ConfigureStyleElement(itemStyle.DropArrow, _
                13, _
                itemStyle.height, _
                itemStyle.width + itemStyle.left - 13, _
                itemStyle.top)
        With .DropArrow
            .Font.Name = FONT_NAME_ICON
            .Caption = itemStyle.ChrDropArrowOff
            .TextAlign = fmTextAlignCenter
            .BackColor = itemStyle.ColorBackGroundOff
            .ForeColor = itemStyle.ColorDropArrowOff
            .BorderStyle = fmBorderStyleNone
            Call SetCommonFontProperties(itemStyle.DropArrow, itemStyle)
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addBtnClear
' Purpose: Adding clear button style for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addBtnClear(ByRef itemStyle As clsModernStyle)
    With itemStyle
        Set .BtnClear = CreateStyledLabel(itemStyle, BTN_CLEAR, 0)
        Call ConfigureStyleElement(itemStyle.BtnClear, _
                10, _
                itemStyle.height, _
                itemStyle.width + itemStyle.left + 2, _
                itemStyle.top)
        With .BtnClear
            .Font.Name = FONT_NAME_ICON
            .Caption = VBA.ChrW$(59719)
            .BackStyle = fmBackStyleTransparent
            .ForeColor = itemStyle.ColorDropArrowOff
            .BorderStyle = fmBorderStyleNone
            .Visible = False
        End With
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addCheckBox
' Purpose: Adding checkbox style for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addCheckBox(ByRef itemStyle As clsModernStyle, ByRef ChrChkBoxBtnOff As String, ByRef ChrChkBoxBtnOn As String)
    With itemStyle

        .control.Font.Size = 8
        If .control.height < 20 Then .control.height = 20

        Set .ChkBoxBtn = .control.Parent.Controls.Add("Forms.Label.1", .control.Name & "_Btn")

        .ChkBoxBtn.left = .control.left
        .ChkBoxBtn.top = .control.top
        .ChkBoxBtn.width = .control.width
        .ChkBoxBtn.height = .control.height
        .ChkBoxBtn.Font.Name = FONT_NAME_ICON
        .ChkBoxBtn.Font.Size = .control.Font.Size * 1.5

        If .control.Value Then
            .ChkBoxBtn.Caption = ChrChkBoxBtnOn
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOn
        Else
            .ChkBoxBtn.Caption = ChrChkBoxBtnOff
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOff
        End If

        '.ChkBoxBtn.BackStyle = fmBackStyleTransparent
        '.ChkBoxBtn.BorderStyle = fmBorderStyleSingle

        If .control.width - 15 > 0 Then

            Set .ChkBoxCaption = .control.Parent.Controls.Add("Forms.Label.1", .control.Name & "_Caption")

            .ChkBoxCaption.left = .control.left + 15
            .ChkBoxCaption.top = .control.top + 1
            .ChkBoxCaption.width = .control.width - 15
            .ChkBoxCaption.height = .control.height
            .ChkBoxCaption.Font.Name = .control.Font.Name
            .ChkBoxCaption.Font.Size = .control.Font.Size
            .ChkBoxCaption.Caption = .control.Caption
            .ChkBoxCaption.BackColor = .control.BackColor
            .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOff
            .ChkBoxCaption.ZOrder (0)

            If .control.Value Then
                .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOn
            Else
                .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOff
            End If

            '.ChkBoxCaption.BackStyle = fmBackStyleTransparent
            '.ChkBoxCaption.BorderStyle = fmBorderStyleSingle
        End If

        If .control.Value Then
            .ChkBoxBtn.Caption = ChrChkBoxBtnOn
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOn
            If .control.width - 15 > 0 Then .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOn
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: addCheckBoxSwitch
' Purpose: Adding switch style for control
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub addCheckBoxSwitch(ByRef itemStyle As clsModernStyle)
    With itemStyle
        .control.Font.Size = 8
        If .control.height < 20 Then .control.height = 20

        Set .TgBorder = .control.Parent.Controls.Add("Forms.Label.1", .Name & "_Border")
        .TgBorder.Visible = .Visible
        .TgBorder.left = .control.left
        .TgBorder.top = .control.top
        .TgBorder.width = .control.width
        If .TgBorder.width < 28 Then .TgBorder.width = 28

        .TgBorder.height = .control.height
        .TgBorder.Font.Name = FONT_NAME_ICON
        .TgBorder.Font.Size = .control.Font.Size * 3.25
        .TgBorder.BackColor = .control.BackColor
        If .control.Value Then
            .TgBorder.ForeColor = .ColorTgBorderOn
            .TgBorder.Caption = VBA.ChrW(ToggleOn)
        Else
            .TgBorder.ForeColor = .ColorTgBorderOff
            .TgBorder.Caption = VBA.ChrW(ToggleOff)
        End If

        '.TgBorder.BackStyle = fmBackStyleTransparent
        '.TgBorder.BorderStyle = fmBorderStyleSingle

        If .control.width - 29 > 0 Then

            Set .ChkBoxCaption = .control.Parent.Controls.Add("Forms.Label.1", .Name & "_Caption")

            .ChkBoxCaption.Visible = .Visible
            .ChkBoxCaption.left = .TgBorder.left + 29
            .ChkBoxCaption.top = .TgBorder.top + 8
            .ChkBoxCaption.width = .control.width - 29
            .ChkBoxCaption.height = .control.height * 0.6
            .ChkBoxCaption.Font.Name = .control.Font.Name
            .ChkBoxCaption.Font.Size = .control.Font.Size
            .ChkBoxCaption.Caption = .control.Caption
            .ChkBoxCaption.BackColor = .control.BackColor
            '.ChkBoxCaption.ForeColor = .Control.ForeColor
            '.ChkBoxCaption.BackStyle = fmBackStyleTransparent
            '.ChkBoxCaption.BorderStyle = fmBorderStyleSingle

            If .control.Value Then
                .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOn
            Else
                .ChkBoxCaption.ForeColor = .ColorChkBoxCaptionOff
            End If
        End If

        Set .ChkBoxBtn = .control.Parent.Controls.Add("Forms.Label.1", .Name & "_Thumb")

        .ChkBoxBtn.Visible = .Visible
        .ChkBoxBtn.top = .control.top + 9.5
        .ChkBoxBtn.width = .TgBorder.width
        .ChkBoxBtn.height = .control.height
        .ChkBoxBtn.Font.Name = FONT_NAME_ICON
        .ChkBoxBtn.Font.Size = .control.Font.Size * 0.9
        .ChkBoxBtn.Caption = VBA.ChrW(ToggleThumb)
        .ChkBoxBtn.BackStyle = fmBackStyleTransparent
        .ChkBoxBtn.BackColor = .control.BackColor
        If .control.Value Then
            .ChkBoxBtn.left = .control.left + 17
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOn
        Else
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOff
            .ChkBoxBtn.left = .left + 2
        End If

        .ChkBoxBtn.ZOrder (0)
        '.TgBorder.BorderStyle = fmBorderStyleSingle

        If .control.Value Then
            .ChkBoxBtn.ForeColor = .ColorChkBoxBtnOn
            .TgBorder.ForeColor = .ColorTgBorderOn
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: HandleExitEvent
' Purpose: Reset style for all controls
'---------------------------------------------------------------------------------------
Private Sub HandleExitEvent()
    Dim item        As clsModernStyle
    For Each item In Me.Items
        Call exitControl(item)
    Next item
End Sub
'---------------------------------------------------------------------------------------
' Procedure: exitControl
' Purpose: Reset control style on focus loss
' Parameters:
'   itemStyle - Style object for configuration
'---------------------------------------------------------------------------------------
Private Sub exitControl(ByRef itemStyle As clsModernStyle)
    With itemStyle
        If Not IsControlActive(itemStyle) Then Exit Sub
        Select Case .ControlType
            Case CONTROL_TYPE_TEXTBOX, CONTROL_TYPE_COMBOBOX, CONTROL_TYPE_LISTBOX:
                With .BarTitle
                    If VBA.Len(itemStyle.control) = 0 Then
                        .ForeColor = itemStyle.ColorBarTitleOff
                        .top = itemStyle.top
                        .left = itemStyle.left + 2
                        .Font.Size = itemStyle.FontSizeTitleOff
                        Call btnClearVisible(itemStyle, False)
                    Else
                        .ForeColor = itemStyle.ColorBarTitleOn
                        .top = itemStyle.top - 17
                        .left = itemStyle.left
                        .Font.Size = itemStyle.FontSizeTitleOn
                        If itemStyle.ControlType = CONTROL_TYPE_LISTBOX Then .ForeColor = itemStyle.ColorBarTitleOff
                        Call btnClearVisible(itemStyle, True)
                    End If
                End With

                With .BarBottom
                    .BackColor = itemStyle.ColorBarBottomOff
                    .height = 1
                End With

                If Not .BackGround Is Nothing Then
                    .BackGround.BackColor = itemStyle.ColorBackGroundOff
                End If

                .BarIcon.ForeColor = itemStyle.ColorBarIconOff

                If Not .DropArrow Is Nothing Then
                    With .DropArrow
                        .ForeColor = itemStyle.ColorDropArrowOff
                        .BackColor = itemStyle.ColorBackGroundOff
                        .Caption = itemStyle.ChrDropArrowOff
                    End With
                End If
        End Select
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: btnClearVisible
' Purpose: Managing visibility of clear button for control
' Parameters:
'   itemStyle - Style object for configuration
'   bVisible - Button visibility flag
'---------------------------------------------------------------------------------------
Private Sub btnClearVisible(ByRef itemStyle As clsModernStyle, ByRef bVisible As Boolean)
    If itemStyle.BtnClear Is Nothing Then Exit Sub
    itemStyle.BtnClear.Visible = bVisible
End Sub
'---------------------------------------------------------------------------------------
' Procedure: HandleEnterEvent
' Purpose: Activating control style on focus gain
'---------------------------------------------------------------------------------------
Private Sub HandleEnterEvent()
    Call HandleExitEvent
    With Me
        If .control Is Nothing Then Exit Sub
        If Not IsControlActive(Me) Then Exit Sub
        .Visible = True
        Select Case .ControlType
            Case CONTROL_TYPE_TEXTBOX, CONTROL_TYPE_LISTBOX, CONTROL_TYPE_COMBOBOX:
                With .BarBottom
                    .BackColor = Me.ColorBarBottomOn
                    .height = 1.2
                End With

                With mBarTitle
                    .ForeColor = Me.ColorBarTitleOn
                    .top = Me.top - 17
                    .Font.Size = Me.FontSizeTitleOn
                    .left = Me.left
                End With

                If Not .BackGround Is Nothing Then
                    .BackGround.BackColor = .ColorBackGroundOn
                End If

                .BarIcon.ForeColor = .ColorBarIconOn

                If Not .DropArrow Is Nothing Then
                    .DropArrow.ForeColor = .ColorDropArrowOn
                    .DropArrow.BackColor = .ColorBackGroundOn
                    Call mDropArrow_Click
                End If
        End Select
    End With
End Sub
'================================================================================
' Section: Control event handling
'================================================================================
' Event handlers for style activation (focusing on element)
Private Sub mListBox_Change()
    Call HandleEnterEvent
End Sub

Private Sub mListBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call HandleEnterEvent
End Sub

Private Sub mListBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleEnterEvent
End Sub

Private Sub mTextBox_Change()
    Call HandleEnterEvent
End Sub

Private Sub mTextbox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call HandleEnterEvent
End Sub

Private Sub mTextbox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleEnterEvent
End Sub

Private Sub mComboBox_Change()
    Call HandleEnterEvent
End Sub

Private Sub mComboBox_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call HandleEnterEvent
End Sub

Private Sub mComboBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call HandleEnterEvent
End Sub

' Event handlers for style reset (focus loss)
Private Sub myForm_Click()
    Call HandleExitEvent
End Sub

Private Sub mFrame_Click()
    Call HandleExitEvent
End Sub

Private Sub mLabel_Click()
    Call HandleExitEvent
End Sub

Private Sub mCommandButton_Click()
    Call HandleExitEvent
End Sub

' Event handlers for specific controls
Private Sub mDropArrow_Click()
    If Not IsControlActive(Me) Then Exit Sub
    With Me
        Select Case .ControlType
            Case CONTROL_TYPE_COMBOBOX:
                .ComboBox.DropDown
                .DropArrow.Caption = .ChrDropArrowOn
        End Select
    End With
End Sub

Private Sub mBtnClear_Click()
    Me.control.Value = vbNullString
End Sub

Private Sub mChkBoxBtn_Click()
    Call UpdateSwitchState(Me.control.Value, True)
End Sub

Private Sub mTgBorder_Click()
    Call UpdateSwitchState(Me.control.Value, True)
End Sub

Private Sub mCheckBox_Change()
    Call UpdateSwitchState(Not Me.control.Value, False)
End Sub

Private Sub mChkBoxCaption_Click()
    Call UpdateSwitchState(Me.control.Value, True)
End Sub

Private Sub mOptionButton_Change()
    Call UpdateSwitchState(Not Me.control.Value, False)
End Sub

'================================================================================
' Section: Internal helper methods
'================================================================================
'---------------------------------------------------------------------------------------
' Procedure: UpdateSwitchState
' Purpose: Internal method to update switch state
' Parameters:
'   Value - New switch value
'   isChangeValue - Flag indicating whether to change control value
'   isCheckBox - Flag indicating whether element is a checkbox
'---------------------------------------------------------------------------------------
Private Sub UpdateSwitchState(ByVal Value As Boolean, ByRef isChangeValue As Boolean)
    With Me.control
        If Not IsControlActive(Me) Then Exit Sub
        If Value Then
            If isChangeValue Then .Value = False
            Call UpdateSwitchVisualState(False)
        Else
            If isChangeValue Then .Value = True
            Call UpdateSwitchVisualState(True)
        End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure: UpdateSwitchVisualState
' Purpose: Internal method to update switch visual state
' Parameters:
'   isEnabled - Flag indicating whether state is enabled
'   isCheckBox - Flag indicating whether element is a checkbox
'---------------------------------------------------------------------------------------
Private Sub UpdateSwitchVisualState(ByVal isEnabled As Boolean)
    ' Update position/symbol of element depending on type and state
    With Me
        Select Case .control.Tag
            Case CONTROL_SWITCH:
                ' Update for toggle switch
                .ChkBoxBtn.left = .control.left + IIf(isEnabled, 17, 2)
                If Not .TgBorder Is Nothing Then
                    .TgBorder.Caption = VBA.ChrW(IIf(isEnabled, ToggleOn, ToggleOff))
                End If
            Case Else
                ' Update for checkbox
                If .ControlType = CONTROL_TYPE_CHECKBOX Then
                    .ChkBoxBtn.Caption = IIf(isEnabled, .ChrChkBoxBtnOn, .ChrChkBoxBtnOff)
                Else
                    .ChkBoxBtn.Caption = IIf(isEnabled, .ChrOptBoxBtnOn, .ChrOptBoxBtnOff)
                End If
        End Select

        ' Update colors
        .ChkBoxBtn.ForeColor = IIf(isEnabled, .ColorChkBoxBtnOn, .ColorChkBoxBtnOff)
        If Not .TgBorder Is Nothing Then
            .TgBorder.ForeColor = IIf(isEnabled, .ColorTgBorderOn, .ColorTgBorderOff)
        End If
        If Not .ChkBoxCaption Is Nothing Then
            .ChkBoxCaption.ForeColor = IIf(isEnabled, .ColorChkBoxCaptionOn, .ColorChkBoxCaptionOff)
        End If
    End With
End Sub
'---------------------------------------------------------------------------------------
' Procedure: IsControlActive
' Purpose: Helper method to check if control is active
' Returns: True if element is visible, unlocked and available
'---------------------------------------------------------------------------------------
Private Function IsControlActive(ByRef itemStyle As clsModernStyle) As Boolean
    With itemStyle
        IsControlActive = .Visible And Not .Locked And .Enabled
    End With
End Function

'---------------------------------------------------------------------------------------
' Procedure: ConfigureStyleElement
' Purpose: Internal method for configuring style element properties
' Parameters:
'   element - Style element to configure
'   itemStyle - Style object
'   width - Element width
'   height - Element height
'   left - Left position
'   top - Top position
'---------------------------------------------------------------------------------------
Private Sub ConfigureStyleElement(ByRef element As MSForms.label, _
        ByVal width As Single, ByVal height As Single, _
        ByVal left As Single, ByVal top As Single)
    With element
        .width = width
        .height = height
        .left = left
        .top = top
    End With
End Sub
