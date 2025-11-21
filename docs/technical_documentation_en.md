# Technical Documentation for VBA Modern Style UserForms Project

## Table of Contents
1. [Class Overview](#class-overview)
2. [Class Architecture](#class-architecture)
3. [Properties](#properties)
4. [Methods](#methods)
5. [Events](#events)
6. [Constants and Enumerations](#constants-and-enumerations)
7. [Implementation Details](#implementation-details)
8. [Dependencies](#dependencies)

## Class Overview

The VBA Modern Style UserForms project is a VBA class library designed to style MSForms controls in Excel. The main class `clsModernStyle` implements modern visual effects, animations, and enhanced visual feedback for custom forms.

### Purpose
The `clsModernStyle` class is designed to apply modern design to MSForms controls with implementations of visual effects such as focus animation, color and font configuration, adding icons and visual elements.

### Main Features
- Applying modern style to various controls (TextBox, ComboBox, ListBox, CheckBox, OptionButton, etc.)
- Supporting focus animation
- Configuring color and font
- Adding icons and visual elements
- Managing visibility and state of controls

## Class Architecture

### Event Handlers
The class uses events to track changes in controls:

```vba
Private WithEvents mUserForm As MSForms.UserForm
Private WithEvents mTextBox As MSForms.TextBox
Private WithEvents mComboBox As MSForms.ComboBox
Private WithEvents mListBox As MSForms.ListBox
Private WithEvents mFrame As MSForms.Frame
Private WithEvents mLabel As MSForms.label
Private WithEvents mCommandButton As MSForms.CommandButton
Private WithEvents mCheckBox As MSForms.CheckBox
Private WithEvents mOptionButton As MSForms.OptionButton
```

### Data Structure
The class stores information about controls in a collection:
- `mStyleItems` - collection of all style items
- `mControl` - main control
- `mControlType` - control type
- `mControlName` - control name
- `mControlTipText` - tooltip text

### Core Class Properties

| Property | Type | Description |
|----------|------|-------------|
| `control` | MSForms.control | The main control being styled |
| `ControlType` | String | The type of control (e.g., "TextBox", "ComboBox") |
| `Name` | String | The name of the control |
| `ControlTipText` | String | The tooltip text for the control |
| `Visible` | Boolean | The visibility of the control and all associated style elements |
| `Locked` | Boolean | The locking state of the control and all associated style elements |
| `Enabled` | Boolean | The enabled state of the control and all associated style elements |
| `top` | Single | The top position of the control |
| `left` | Single | The left position of the control |
| `height` | Single | The height of the control |
| `width` | Single | The width of the control |
| `FontSizeTitleOff` | Integer | The font size for inactive state |
| `FontSizeTitleOn` | Integer | The font size for active state |
| `FontName` | String | The font name for the control |

## Properties

### Main Control Properties
- `control` - Gets or sets the main control that will be styled
- `ControlType` - Gets or sets the type of control (e.g., "TextBox", "ComboBox")
- `Name` - Gets or sets the name of the control
- `ControlTipText` - Gets or sets the tooltip text for the control

### Visibility and State Properties
- `Visible` - Gets or sets the visibility of the control and all associated style elements
- `Locked` - Gets or sets the locking state of the control and all associated style elements
- `Enabled` - Gets or sets the enabled state of the control and all associated style elements

### Positioning and Size Properties
- `top` - Gets or sets the top position of the control
- `left` - Gets or sets the left position of the control
- `height` - Gets or sets the height of the control
- `width` - Gets or sets the width of the control

### Font Properties
- `FontSizeTitleOff` - Gets or sets the font size for inactive state
- `FontSizeTitleOn` - Gets or sets the font size for active state
- `FontName` - Gets or sets the font name for the control

### Color Properties
- `ColorBarTitleOn` - Gets or sets the title color in active state
- `ColorBarTitleOff` - Gets or sets the title color in inactive state
- `ColorBarBottomOn` - Gets or sets the bottom line color in active state
- `ColorBarBottomOff` - Gets or sets the bottom line color in inactive state
- `ColorBackGroundOn` - Gets or sets the background color in active state
- `ColorBackGroundOff` - Gets or sets the background color in inactive state
- `ColorBarIconOn` - Gets or sets the icon color in active state
- `ColorBarIconOff` - Gets or sets the icon color in inactive state
- `ColorDropArrowOn` - Gets or sets the dropdown arrow color in active state
- `ColorDropArrowOff` - Gets or sets the dropdown arrow color in inactive state
- `ColorTgBorderOn` - Gets or sets the toggle border color in active state
- `ColorTgBorderOff` - Gets or sets the toggle border color inactive state
- `ColorChkBoxBtnOn` - Gets or sets the checkbox button color in active state
- `ColorChkBoxBtnOff` - Gets or sets the checkbox button color in inactive state
- `ColorChkBoxCaptionOn` - Gets or sets the checkbox caption color in active state
- `ColorChkBoxCaptionOff` - Gets or sets the checkbox caption color in inactive state

### Character Properties
- `ChrDropArrowOn` - Gets or sets the dropdown arrow character in active state
- `ChrDropArrowOff` - Gets or sets the dropdown arrow character in inactive state
- `ChrChkBoxBtnOn` - Gets or sets the checkbox button character in active state
- `ChrChkBoxBtnOff` - Gets or sets the checkbox button character in inactive state
- `ChrOptBoxBtnOn` - Gets or sets the option button character in active state
- `ChrOptBoxBtnOff` - Gets or sets the option button character in inactive state

### Additional Element Properties
- `BarBottom` - Gets or sets the bottom line of the control
- `BarTitle` - Gets or sets the title of the control
- `BarIcon` - Gets or sets the icon of the control
- `BackGround` - Gets or sets the background of the control
- `DropArrow` - Gets or sets the dropdown arrow
- `BtnClear` - Gets or sets the clear button
- `TgBorder` - Gets or sets the toggle border
- `ChkBoxBtn` - Gets or sets the checkbox button
- `ChkBoxCaption` - Gets or sets the checkbox caption

### Collection Properties
- `StyleItems` - Gets or sets the collection of all style items
- `Count` - Gets the number of items in the collection
- `getItemByIndex` - Gets an item from the collection by index
- `getItemByName` - Gets an item from the collection by name
- `Version` - Gets version information about the class

## Methods

### Main Initialization Methods
- `Initialize` - Initializes style for all form controls

**Syntax:**
```vba
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
```

**Parameters:**
- `Form` - reference to UserForm to which style is applied
- `ColorBarTitleOn` - title color in active state (default 14854934)
- `ColorBarTitleOff` - title color in inactive state (default 10395294)
- `ColorBarBottomOn` - bottom line color in active state (default 14854934)
- `ColorBarBottomOff` - bottom line color in inactive state (default 10395294)
- `ColorBackGroundOn` - background color in active state (default vbWhite)
- `ColorBackGroundOff` - background color inactive state (default 1647476)
- `ColorBarIconOn` - icon color in active state (default 14854934)
- `ColorBarIconOff` - icon color inactive state (default 10395294)
- `ColorDropArrowOn` - dropdown arrow color in active state (default vbBlack)
- `ColorDropArrowOff` - dropdown arrow color in inactive state (default 10395294)
- `ColorTgBorderOn` - toggle border color in active state (default 14854934)
- `ColorTgBorderOff` - toggle border color in inactive state (default 10395294)
- `ColorChkBoxBtnOn` - checkbox button color in active state (default vbBlack)
- `ColorChkBoxBtnOff` - checkbox button color inactive state (default 10395294)
- `ChrDropArrowOn` - dropdown arrow character in active state (default ArrowOn)
- `ChrDropArrowOff` - dropdown arrow character in inactive state (default ArrowOff)
- `ColorChkBoxCaptionOn` - checkbox caption color in active state (default 14854934)
- `ColorChkBoxCaptionOff` - checkbox caption color in inactive state (default 10395294)
- `ChrChkBoxBtnOn` - checkbox button character in active state (default CheckboxComposite)
- `ChrChkBoxBtnOff` - checkbox button character in inactive state (default CheckBox1)
- `ChrOptBoxBtnOn` - option button character in active state (default CircleFill)
- `ChrOptBoxBtnOff` - option button character in inactive state (default CircleRing)

### Control Styling Methods
- `ApplyControlStyle` - applies style depending on control type
- `SetCommonStyleProperties` - sets common style properties for a control
- `setTextBoxStyle` - sets style for text box
- `setComboBoxStyle` - sets style for combo box
- `setListBoxStyle` - sets style for list box

### Style Element Addition Methods
- `CreateStyledLabel` - creates and configures main properties of additional element
- `SetCommonFontProperties` - sets common font properties for a control
- `addBarBottom` - adds bottom style line for control
- `addBarTitle` - adds style title for control
- `addBarIcon` - adds style icon for control
- `addBackGround` - adds style background for control
- `addDropArrow` - adds dropdown arrow style for control
- `addBtnClear` - adds clear button style for control
- `addCheckBox` - adds checkbox style for control
- `addCheckBoxSwitch` - adds toggle switch style for control

### Event Handling Methods
- `HandleExitEvent` - reset style for all controls
- `exitControl` - reset control style on focus loss
- `btnClearVisible` - managing visibility of clear button for control
- `HandleEnterEvent` - activating control style on focus gain

## Events

### Text Box Events
- `mTextBox_Change` - text change event in text box
- `mTextbox_MouseDown` - mouse click event on text box
- `mTextbox_KeyUp` - key release event when focused on text box

### Combo Box Events
- `mComboBox_Change` - value change event in combo box
- `mComboBox_KeyUp` - key release event when focused on combo box
- `mComboBox_MouseDown` - mouse click event on combo box

### List Box Events
- `mListBox_Change` - value change event in list box
- `mListBox_MouseDown` - mouse click event on list box
- `mListBox_KeyUp` - key release event when focused on list box

### Other Control Events
- `mUserForm_Click` - click event on user form
- `mFrame_Click` - click event on frame
- `mLabel_Click` - click event on label
- `mCommandButton_Click` - click event on command button

### Specific Element Events
- `mDropArrow_Click` - click event on dropdown arrow
- `mBtnClear_Click` - click event on clear button
- `mChkBoxBtn_Click` - click event on checkbox button
- `mTgBorder_Click` - click event on toggle border
- `mCheckBox_Change` - checkbox state change event
- `mChkBoxCaption_Click` - click event on checkbox caption
- `mOptionButton_Change` - option button state change event

## Constants and Enumerations

### Icon Enumeration
```vba
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
    FavoriteStar = 5918                    ' Star (normal)
    FavoriteStarFill = 59189                ' Star (filled)
    Heart = 60241                           ' Heart (normal)
    HeartFill = 60242                       ' Heart (filled)
    InkingColorFill = 60775                 ' Brush (filled)
    InkingColorOutline = 6074              ' Brush (outline)
    PaginationDotOutline10 = 61734          ' Dot (outline)
    PaginationDotSolid10 = 61735            ' Dot (filled)
    PasswordChar = 149                      ' Character for hiding password
    RadioBtnOff = 60618                     ' Radio button (off)
    RadioBtnOn = 60619                      ' Radio button (on)
    ToggleOff = 60434                       ' Toggle switch (off)
    ToggleOn = 60433                        ' Toggle switch (on)
    ToggleThumb = 60436                     ' Toggle switch thumb
End Enum
```

### Constants
```vba
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
```

## Implementation Details

### Helper Methods
- `UpdateSwitchState` - internal method to update switch state
- `UpdateSwitchVisualState` - internal method to update switch visual state
- `IsControlActive` - helper method to check if control is active
- `ConfigureStyleElement` - internal method for configuring style element properties
- `IsControlInCollection` - check if control already exists in the collection
- `SetControlEnabled` - internal method to set the availability state of a control
- `SetControlVisibility` - internal method to set the visibility of a control
- `SetControlLock` - internal method to set the locking state of a control

### Cleanup Method
- `Class_Terminate` - clean up objects when class is terminated

## Dependencies

- Microsoft Forms 2.0 Object Library
- VBA runtime environment