# User Guide for VBA Modern Style UserForms Project

## Table of Contents
1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation and Setup](#installation-and-setup)
4. [Quick Start](#quick-start)
5. [Main Features](#main-features)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Creating Toggle Switches](#creating-toggle-switches)
9. [Adding Icons](#adding-icons)
10. [Working with Style Collection](#working-with-style-collection)
11. [Troubleshooting](#troubleshooting)
12. [Frequently Asked Questions](#frequently-asked-questions)

## Introduction

The VBA Modern Style UserForms project provides modern styling for MSForms controls in Excel. This tool allows you to easily enhance the appearance of user forms with minimal effort, adding contemporary visual elements and animations.

### What This Project Can Do:
- Apply modern design to various controls (TextBox, ComboBox, ListBox, CheckBox, OptionButton, etc.)
- Provide focus animation for improved user experience
- Configure colors and fonts for controls
- Add icons and visual elements
- Manage visibility and state of controls

## System Requirements

- Microsoft Excel (2010 or newer recommended)
- VBA support enabled
- Microsoft Forms 2.0 Object Library
- Windows 7 or newer

## Installation and Setup

### Step 1: Import the Class
1. Open Excel and go to the VBA editor (press Alt+F11)
2. In the menu, select "File" > "Import File"
3. Select the `clsModernStyle.cls` file from the `vba-files/Class/` directory
4. Click "Open" to import the class

### Step 2: Configure References
1. In the VBA editor, select "Tools" > "References"
2. Find and check the box next to "Microsoft Forms 2.0 Object Library"
3. Click "OK" to save changes

### Step 3: Create a User Form
1. In the VBA editor, create a new user form
2. Add controls that you want to style
3. Add a class variable to the form:
```vba
Dim MStyleItem As clsModernStyle
```

## Quick Start

### Simple Usage Example
1. Create a new user form in Excel
2. Add several controls (TextBox, ComboBox, CheckBox)
3. In the properties window of each control, set the `ControlTipText` property with a description
4. In the `UserForm_Initialize` event, add the following code:
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```
5. Run the form to see the modern styling of controls

### Example with Color Configuration
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    MStyleItem.Initialize Me, _
        ColorBarTitleOn:=RGB(0, 100, 200), _
        ColorBarTitleOff:=RGB(120, 120, 120), _
        ColorBarBottomOn:=RGB(0, 100, 200), _
        ColorBarBottomOff:=RGB(180, 180, 180), _
        ColorBackGroundOn:=RGB(255, 255, 255), _
        ColorBackGroundOff:=RGB(245, 245, 245)
End Sub
```

## Main Features

### Style Initialization
The `Initialize` method is the main way to apply styles to all controls on a form:
- Automatically recognizes all supported controls
- Applies modern styles to all elements
- Creates additional visual elements (lines, titles, icons)
- Sets up event handlers for focus animation

### Supported Controls
- TextBox - text fields with title animation and clear button
- ComboBox - combo boxes with animation and arrow
- ListBox - lists with enhanced visual styling
- CheckBox - checkboxes with modern styling
- OptionButton - option buttons with circular icons
- Frame - frames (no style changes)
- Label - labels (no style changes)
- CommandButton - buttons (no style changes)

## Working with Controls

### Text Boxes
For text boxes, the class automatically:
- Applies transparent background and flat border style
- Adds a bottom line that changes color on focus
- Creates an animated title that moves when text is entered
- Adds an icon on the left (if specified in the Tag property)
- Provides a clear button (appears when text is entered)

### Combo Boxes
For combo boxes, the class:
- Hides the standard dropdown arrow
- Adds a custom arrow with animation
- Applies transparent background
- Provides focus animation similar to text boxes
- Adds a clear button when needed

### Lists
For lists, the class:
- Applies border style
- Adds bottom line and title
- Provides animation when selecting items
- Supports scrolling and multiple item selection

### Checkboxes and Option Buttons
For checkboxes and option buttons, the class:
- Replaces standard controls with styled icons
- Provides visual feedback when changing state
- Supports different styles for active and inactive states
- Allows creating modern toggle switches (see "Creating Toggle Switches" section)

## Style Configuration

### Color Schemes
The class provides extensive color configuration options:
- Title colors (active/inactive state)
- Bottom line colors (active/inactive state)
- Background colors (active/inactive state)
- Icon colors (active/inactive state)
- Dropdown arrow colors
- Toggle border colors
- Checkbox button colors

### Configuration via Initialize Method
Colors can be configured during initialization:
```vba
MStyleItem.Initialize Me, _
    ColorBarTitleOn:=RGB(255, 0, 0), _
    ColorBarTitleOff:=RGB(128, 128, 128), _
    ColorBarBottomOn:=RGB(0, 0, 255), _
    ColorBarBottomOff:=RGB(192, 192, 192)
```

### Dynamic Configuration
After initialization, you can change colors for individual elements:
```vba
With MStyleItem.getItemByName(TextBox1.Name)
    .ColorBackGroundOff = vbRed
    .ColorBackGroundOn = vbGreen
End With
```

### Font Configuration
The class automatically:
- Sets the font name (default Segoe UI)
- Configures font sizes for active and inactive states
- Supports font size configuration via `FontSizeTitleOff` and `FontSizeTitleOn` properties

## Creating Toggle Switches

To create a modern toggle switch instead of a regular checkbox:
1. In the form designer, set the `Tag` property of the checkbox to "SWITCH"
2. Or programmatically:
```vba
MyCheckBox.Tag = "SWITCH"
```
3. When initializing styles, the checkbox will be displayed as a modern toggle switch with animation

### Toggle Switch Types
- Regular toggle - circular button with animation
- Styled toggle - using icons from the enumIcons enumeration
- Custom toggle - with configurable colors and sizes

## Adding Icons

### Using Built-in Icons
The class includes an `enumIcons` enumeration with various icons:
- ArrowOff, ArrowOn (for dropdowns)
- CheckBox1, CheckboxComposite, CheckboxFill
- CheckMark
- CircleFill, CircleRing (for option buttons)
- FavoriteStar, FavoriteStarFill
- Heart, HeartFill
- PasswordChar
- RadioBtnOff, RadioBtnOn
- ToggleOff, ToggleOn, ToggleThumb

### Setting Icons
To add an icon to a control:
1. In the form designer, set the `Tag` property of the control to the numeric value of the icon
2. Or programmatically:
```vba
MyTextBox.Tag = 59193  ' Using numeric value for icon
' Or
MyTextBox.Tag = "61735"  ' Using string value for icon
```
3. The icon will appear to the left of the control

### Icon Configuration
- Icons are displayed using the Segoe MDL2 Assets font
- Icon color changes depending on the control state
- Icons automatically scale to the control size

## Working with Style Collection

### Accessing Individual Elements
After initialization, all styled controls are stored in a collection:
```vba
' Getting element by name
Dim item As clsModernStyle
Set item = MStyleItem.getItemByName(TextBox1.Name)

' Getting element by index
Set item = MStyleItem.getItemByIndex(1)
```

### Getting the Number of Elements
```vba
Dim count As Byte
count = MStyleItem.Count
```

### Iterating Through All Elements
```vba
Dim item As clsModernStyle
For Each item In MStyleItem.StyleItems
    ' Processing each element
    Debug.Print item.Name
Next item
```

### Changing Properties of Individual Elements
```vba
' Changing background color of a specific element
With MStyleItem.getItemByName(TextBox1.Name)
    .ColorBackGroundOn = RGB(255, 255, 200)
    .ColorBackGroundOff = RGB(255, 255, 255)
End With
```

## Troubleshooting

### Display Issues
- Ensure Microsoft Forms 2.0 Object Library is enabled in references
- Check that all controls are added before calling the Initialize method
- Ensure the MultiUse property is set to True for the class

### Animation Issues
- Check that control events are not overloaded with other handlers
- Ensure control properties are not changed manually while the class is running
- Verify that the class is not initialized multiple times

### Performance Issues
- Reduce the number of controls on the form
- Avoid frequent calls to methods for getting elements from the collection
- Use visibility and availability properties instead of removing elements

### Common Errors
- "Object variable not set" - ensure the class variable is properly initialized
- "Method or data member not found" - check that the class is properly imported
- "Can't assign to property" - avoid direct assignment to nested objects without checking for Nothing

## Frequently Asked Questions

### Question: How to change colors after initialization?
**Answer:** Use the `getItemByName` or `getItemByIndex` methods to get a specific element and change its color properties.

### Question: Are all controls supported?
**Answer:** The class supports main MSForms controls. Some controls (e.g., ScrollBar, SpinButton) are supported partially.

### Question: Can multiple class instances be used?
**Answer:** Yes, you can create multiple class instances for different forms, but it's not recommended to use multiple instances for one form.

### Question: How to add custom icons?
**Answer:** The class uses the Segoe MDL2 Assets font, which contains many built-in icons. For custom icons, you can use images or characters from other fonts.

### Question: Is the class compatible with different Excel versions?
**Answer:** The class is tested with Excel 2010 and newer. Compatibility with earlier versions is not guaranteed.

### Question: Can animation be configured?
**Answer:** The current version does not provide direct animation configuration, but you can change visual effects through color properties and font sizes.

### Question: How to handle events of styled elements?
**Answer:** Events of original controls continue to work as usual. The class only adds visual effects and does not change event handling logic.