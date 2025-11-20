# Technical Documentation for VBA Modern Style UserForms Project

## Table of Contents
1. [Project Overview](#project-overview)
2. [Class Architecture clsModernStyle](#class-architecture-clsmodernstyle)
3. [Class Properties Detail](#class-properties-detail)
4. [Class Methods Detail](#class-methods-detail)
5. [Events and Event Handling](#events-and-event-handling)
6. [Enumerations and Constants](#enumerations-and-constants)
7. [Internal Structure and Helper Methods](#internal-structure-and-helper-methods)

## Project Overview

The VBA Modern Style UserForms project is a VBA class library designed to style MSForms controls in Excel. The main class `clsModernStyle` implements modern visual effects, animations, and enhanced visual feedback for custom forms.

### Purpose
The `clsModernStyle` class is designed to apply modern design to MSForms controls with implementations of visual effects such as focus animation, color and font configuration, adding icons and visual elements.

### Main Features
- Applying modern style to various controls (TextBox, ComboBox, ListBox, CheckBox, OptionButton, etc.)
- Supporting focus animation
- Configuring color and font
- Adding icons and visual elements
- Managing visibility and state of controls

## Class Architecture clsModernStyle

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

## Class Properties Detail

### Main Control Properties
- `control` - the main control being styled
  - Type: `MSForms.control`
  - Purpose: Getting or setting the main control that will be styled

- `ControlType` - the type of control
  - Type: `String`
  - Purpose: Getting or setting the type of control (e.g., "TextBox", "ComboBox")

- `Name` - the name of the control
  - Type: `String`
 - Purpose: Getting or setting the name of the control

- `ControlTipText` - tooltip text
  - Type: `String`
  - Purpose: Getting or setting the tooltip text for the control

### Visibility and State Properties
- `Visible` - visibility of the control
  - Type: `Boolean`
  - Purpose: Getting or setting the visibility of the control and all associated style elements

- `Locked` - locking state
  - Type: `Boolean`
  - Purpose: Getting or setting the locking state of the control and all associated style elements

- `Enabled` - enabled state
 - Type: `Boolean`
  - Purpose: Getting or setting the enabled state of the control and all associated style elements

### Positioning and Size Properties
- `top` - top position
  - Type: `Single`
  - Purpose: Getting or setting the top position of the control

- `left` - left position
 - Type: `Single`
  - Purpose: Getting or setting the left position of the control

- `height` - height
  - Type: `Single`
  - Purpose: Getting or setting the height of the control

- `width` - width
  - Type: `Single`
  - Purpose: Getting or setting the width of the control

### Font Properties
- `FontSizeTitleOff` - font size for inactive state
  - Type: `Integer`
  - Purpose: Getting or setting the font size for inactive state

- `FontSizeTitleOn` - font size for active state
  - Type: `Integer`
  - Purpose: Getting or setting the font size for active state

- `FontName` - font name
 - Type: `String`
  - Purpose: Getting or setting the font name for the control

### Color Properties
- `ColorBarTitleOn` - title color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the title color in active state

- `ColorBarTitleOff` - title color in inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the title color in inactive state

- `ColorBarBottomOn` - bottom line color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the bottom line color in active state

- `ColorBarBottomOff` - bottom line color in inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the bottom line color in inactive state

- `ColorBackGroundOn` - background color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the background color in active state

- `ColorBackGroundOff` - background color in inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the background color in inactive state

- `ColorBarIconOn` - icon color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the icon color in active state

- `ColorBarIconOff` - icon color in inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the icon color in inactive state

- `ColorDropArrowOn` - dropdown arrow color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the dropdown arrow color in active state

- `ColorDropArrowOff` - dropdown arrow color in inactive state
 - Type: `XlRgbColor`
  - Purpose: Getting or setting the dropdown arrow color in inactive state

- `ColorTgBorderOn` - toggle border color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the toggle border color in active state

- `ColorTgBorderOff` - toggle border color inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the toggle border color in inactive state

- `ColorChkBoxBtnOn` - checkbox button color in active state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the checkbox button color in active state

- `ColorChkBoxBtnOff` - checkbox button color in inactive state
  - Type: `XlRgbColor`
  - Purpose: Getting or setting the checkbox button color in inactive state

- `ColorChkBoxCaptionOn` - checkbox caption color in active state
  - Type: `Long`
  - Purpose: Getting or setting the checkbox caption color in active state

- `ColorChkBoxCaptionOff` - checkbox caption color in inactive state
  - Type: `Long`
  - Purpose: Getting or setting the checkbox caption color in inactive state

### Character Properties
- `ChrDropArrowOn` - dropdown arrow character in active state
  - Type: `String`
  - Purpose: Getting or setting the dropdown arrow character in active state

- `ChrDropArrowOff` - dropdown arrow character in inactive state
  - Type: `String`
 - Purpose: Getting or setting the dropdown arrow character in inactive state

- `ChrChkBoxBtnOn` - checkbox button character in active state
  - Type: `String`
  - Purpose: Getting or setting the checkbox button character in active state

- `ChrChkBoxBtnOff` - checkbox button character in inactive state
  - Type: `String`
  - Purpose: Getting or setting the checkbox button character in inactive state

- `ChrOptBoxBtnOn` - option button character in active state
 - Type: `String`
  - Purpose: Getting or setting the option button character in active state

- `ChrOptBoxBtnOff` - option button character in inactive state
  - Type: `String`
  - Purpose: Getting or setting the option button character in inactive state

### Additional Element Properties
- `BarBottom` - bottom line of the control
  - Type: `MSForms.label`
  - Purpose: Getting or setting the bottom line of the control

- `BarTitle` - title of the control
 - Type: `MSForms.label`
 - Purpose: Getting or setting the title of the control

- `BarIcon` - icon of the control
  - Type: `MSForms.label`
  - Purpose: Getting or setting the icon of the control

- `BackGround` - background of the control
 - Type: `MSForms.label`
 - Purpose: Getting or setting the background of the control

- `DropArrow` - dropdown arrow
  - Type: `MSForms.label`
  - Purpose: Getting or setting the dropdown arrow

- `BtnClear` - clear button
 - Type: `MSForms.label`
  - Purpose: Getting or setting the clear button

- `TgBorder` - toggle border
  - Type: `MSForms.label`
  - Purpose: Getting or setting the toggle border

- `ChkBoxBtn` - checkbox button
 - Type: `MSForms.label`
  - Purpose: Getting or setting the checkbox button

- `ChkBoxCaption` - checkbox caption
 - Type: `MSForms.label`
  - Purpose: Getting or setting the checkbox caption

### Collection Properties
- `StyleItems` - collection of all style items
  - Type: `Collection`
  - Purpose: Getting or setting the collection of all style items

- `Count` - number of items in the collection
  - Type: `Byte`
 - Purpose: Getting the number of items in the collection

- `getItemByIndex` - getting an item from the collection by index
  - Type: `clsModernStyle`
  - Purpose: Getting an item from the collection by index

- `getItemByName` - getting an item from the collection by name
  - Type: `clsModernStyle`
  - Purpose: Getting an item from the collection by name

- `Version` - version information about the class
  - Type: `String`
  - Purpose: Getting version information about the class

## Class Methods Detail

### Main Initialization Methods
- `Initialize` - initializes style for all form controls
  - Parameters:
    - `Form` - reference to UserForm to which style is applied
    - `ColorBarTitleOn` - title color in active state (default 14854934)
    - `ColorBarTitleOff` - title color in inactive state (default 10395294)
    - `ColorBarBottomOn` - bottom line color in active state (default 14854934)
    - `ColorBarBottomOff` - bottom line color in inactive state (default 10395294)
    - `ColorBackGroundOn` - background color in active state (default vbWhite)
    - `ColorBackGroundOff` - background color inactive state (default 16447476)
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
  - Purpose: Initializes style for all controls on the form

### Control Styling Methods
- `ApplyControlStyle` - applies style depending on control type
  - Parameters:
    - `itemStyle` - style object for configuration
  - Purpose: Applies style depending on control type

- `SetCommonStyleProperties` - sets common style properties for a control
  - Parameters:
    - `itemStyle` - style object for configuration
  - Purpose: Setting common style properties for a control

- `setTextBoxStyle` - sets style for text box
  - Parameters:
    - `itemStyle` - style object for configuration
    - `sPasswordChar` - character for displaying password field (optional parameter)
  - Purpose: Setting style for text box

- `setComboBoxStyle` - sets style for combo box
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Setting style for combo box

- `setListBoxStyle` - sets style for list box
 - Parameters:
    - `itemStyle` - style object for configuration
  - Purpose: Setting style for list box

### Style Element Addition Methods
- `CreateStyledLabel` - creating and configuring main properties of additional element
  - Parameters:
    - `itemStyle` - style object for configuration
    - `controlName` - name of control
    - `zIndex` - element placement order (default 1)
  - Returns: New `MSForms.label` control
  - Purpose: Creating and configuring main properties of additional element

- `SetCommonFontProperties` - setting common font properties for a control
  - Parameters:
    - `label` - `MSForms.label` control
    - `itemStyle` - style object for configuration
  - Purpose: Setting common font properties for a control

- `addBarBottom` - adding bottom style line for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding bottom style line for control

- `addBarTitle` - adding style title for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding style title for control

- `addBarIcon` - adding style icon for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding style icon for control

- `addBackGround` - adding style background for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding style background for control

- `addDropArrow` - adding dropdown arrow style for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding dropdown arrow style for control

- `addBtnClear` - adding clear button style for control
  - Parameters:
    - `itemStyle` - style object for configuration
  - Purpose: Adding clear button style for control

- `addCheckBox` - adding checkbox style for control
  - Parameters:
    - `itemStyle` - style object for configuration
    - `ChrChkBoxBtnOff` - checkbox button character in inactive state
    - `ChrChkBoxBtnOn` - checkbox button character in active state
  - Purpose: Adding checkbox style for control

- `addCheckBoxSwitch` - adding toggle switch style for control
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Adding toggle switch style for control

### Event Handling Methods
- `HandleExitEvent` - reset style for all controls
  - Purpose: Reset style for all controls

- `exitControl` - reset control style on focus loss
  - Parameters:
    - `itemStyle` - style object for configuration
 - Purpose: Reset control style on focus loss

- `btnClearVisible` - managing visibility of clear button for control
  - Parameters:
    - `itemStyle` - style object for configuration
    - `bVisible` - button visibility flag
  - Purpose: Managing visibility of clear button for control

- `HandleEnterEvent` - activating control style on focus gain
 - Purpose: Activating control style on focus gain

## Events and Event Handling

The `clsModernStyle` class uses events to track changes in controls and activate appropriate actions:

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

## Enumerations and Constants

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
    FavoriteStar = 59188                    ' Star (normal)
    FavoriteStarFill = 59189                ' Star (filled)
    Heart = 60241                           ' Heart (normal)
    HeartFill = 60242                       ' Heart (filled)
    InkingColorFill = 60775                 ' Brush (filled)
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

## Internal Structure and Helper Methods

### Helper Methods
- `UpdateSwitchState` - internal method to update switch state
  - Parameters:
    - `Value` - new switch value
    - `isChangeValue` - flag indicating whether to change control value
  - Purpose: Internal method to update switch state

- `UpdateSwitchVisualState` - internal method to update switch visual state
  - Parameters:
    - `isEnabled` - flag indicating whether state is enabled
  - Purpose: Internal method to update switch visual state

- `IsControlActive` - helper method to check if control is active
  - Parameters:
    - `itemStyle` - style object for checking
 - Returns: `True` if element is visible, unlocked and available
 - Purpose: Helper method to check if control is active

- `ConfigureStyleElement` - internal method for configuring style element properties
 - Parameters:
    - `element` - style element to configure
    - `width` - element width
    - `height` - element height
    - `left` - left position
    - `top` - top position
  - Purpose: Internal method for configuring style element properties

- `IsControlInCollection` - check if control already exists in the collection
  - Parameters:
    - `controlName` - name of the control to check
  - Returns: `True` if control exists in collection, `False` otherwise
 - Purpose: Check if control already exists in the collection

- `SetControlEnabled` - internal method to set the availability state of a control
  - Parameters:
    - `control` - control to change availability state
    - `isEnabled` - flag for availability state
  - Purpose: Internal method to set the availability state of a control

- `SetControlVisibility` - internal method to set the visibility of a control
  - Parameters:
    - `control` - control to change visibility
    - `isVisible` - visibility flag
  - Purpose: Internal method to set the visibility of a control

- `SetControlLock` - internal method to set the locking state of a control
  - Parameters:
    - `control` - control to change locking state
    - `isLocked` - flag for locking state
  - Purpose: Internal method to set the locking state of a control

### Cleanup Method
- `Class_Terminate` - clean up objects when class is terminated
  - Purpose: Clean up event handler objects and style elements when class is terminated