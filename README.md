# VBA Modern Style Class

![User Form Example](User_Forms.png)

A VBA class library that provides modern styling for MSForms controls in Excel applications. This class enhances the appearance of UserForms with contemporary design elements, animations, and improved visual feedback.

## Features

- **Modern Design**: Apply sleek, contemporary styling to various MSForms controls
- **Focus Animation**: Visual feedback when controls gain focus with animated elements
- **Color Customization**: Configure colors for different control states (active/inactive)
- **Icon Support**: Add icons to controls using the Segoe MDL2 Assets font
- **Clear Buttons**: Automatic clear buttons for textboxes and combo boxes
- **Toggle Switches**: Modern toggle switch styling for checkboxes and option buttons
- **Responsive Labels**: Labels that animate when controls receive input

## Supported Controls

- TextBox
- ComboBox
- ListBox
- CheckBox
- OptionButton
- Frame
- Label
- CommandButton

## Installation

1. Download the `clsModernStyle.cls` file from the `vba-files/Class/` directory
2. Import the class into your VBA project
3. Ensure you have the Microsoft Forms 2.0 Object Library referenced in your project

## Usage

```vba
' Create an instance of clsModernStyle class
Dim style As New clsModernStyle

' Initialize the styling for your UserForm
style.Initialize Me ' where Me is the UserForm

' The class automatically applies modern styling to all compatible controls on the form
```

### Advanced Usage with Custom Colors

```vba
' Initialize with custom colors
Dim style As New clsModernStyle
style.Initialize Me, _
    ColorBarTitleOn:=RGB(0, 100, 200), _
    ColorBarTitleOff:=RGB(120, 120, 120), _
    ColorBarBottomOn:=RGB(0, 100, 200), _
    ColorBarBottomOff:=RGB(180, 180, 180), _
    ColorBackGroundOn:=RGB(255, 255, 255), _
    ColorBackGroundOff:=RGB(245, 245, 245)
```

### Creating Toggle Switches

To create a toggle switch instead of a regular checkbox, set the Tag property of the control to "SWITCH":

```vba
' In the UserForm designer, set the Tag property of a CheckBox to "SWITCH"
' Or programmatically:
MyCheckBox.Tag = "SWITCH"
```

## Icons

The class includes an enumeration of icons that can be used with controls:

- ArrowOff, ArrowOn (for dropdowns)
- CheckBox1, CheckboxComposite, CheckboxFill
- CheckMark
- CircleFill, CircleRing (for option buttons)
- FavoriteStar, FavoriteStarFill
- Heart, HeartFill
- PasswordChar
- RadioBtnOff, RadioBtnOn
- ToggleOff, ToggleOn, ToggleThumb

## Customization Options

The class allows customization of:

- Font sizes for active and inactive states
- Colors for various control elements
- Character sets for different control types
- Visibility and enabled state of controls

## Version

- Version: 1.0.8
- Creation Date: 10.10.2025 15:30
- Update Date: 01.11.2025 09:59
- Author: VBATools

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.

## Dependencies

- Microsoft Forms 2.0 Object Library
- VBA (Visual Basic for Applications)

## Examples

The repository includes a test form (`frmTest.frm`) demonstrating the capabilities of the class. You can use this as a reference for implementing the styling in your own projects.