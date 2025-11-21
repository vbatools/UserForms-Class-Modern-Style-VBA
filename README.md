# VBA Modern Style Class

![Project Demo](User_Forms.gif)

This repository contains a VBA class library implementation that provides modern styling for MSForms controls in Excel applications. The class enhances the appearance of UserForms with contemporary design elements, animations, and improved visual feedback.

## Table of Contents
1. [Features](#features)
2. [Components](#components)
3. [Installation](#installation)
4. [Quick Start](#quick-start)
5. [Main Functions](#main-functions)
6. [Working with Controls](#working-with-controls)
7. [Style Configuration](#style-configuration)
8. [Troubleshooting](#troubleshooting)

## Features

- **Modern Design**: Apply sleek, contemporary styling to various MSForms controls
- **Focus Animation**: Visual feedback when controls gain focus with animated elements
- **Color Customization**: Configure colors for different control states (active/inactive)
- **Icon Support**: Add icons to controls using the Segoe MDL2 Assets font
- **Clear Buttons**: Automatic clear buttons for textboxes and combo boxes
- **Toggle Switches**: Modern toggle switch styling for checkboxes and option buttons
- **Responsive Labels**: Labels that animate when controls receive input

## Components

- `clsModernStyle.cls`: The main class implementation
- `frmTest.frm`: Test form demonstrating usage
- `modShowForms.bas`: Module containing form display functions
- Documentation in the `docs/` folder:
  - [`docs/technical_documentation_en.md`](docs/technical_documentation_en.md) - Technical documentation in English
 - [`docs/technical_documentation_ru.md`](docs/technical_documentation_ru.md) - Technical documentation in Russian
 - [`docs/user_guide_en.md`](docs/user_guide_en.md) - User guide in English
  - [`docs/user_guide_ru.md`](docs/user_guide_ru.md) - User guide in Russian
 - [`docs/implementation_examples_en.md`](docs/implementation_examples_en.md) - Implementation examples in English
  - [`docs/implementation_examples_ru.md`](docs/implementation_examples_ru.md) - Implementation examples in Russian

## Installation

1. Download the `clsModernStyle.cls` file from the `vba-files/Class/` directory
2. Import the class into your VBA project
3. Ensure you have the Microsoft Forms 2.0 Object Library referenced in your project

## Quick Start

### Simple Usage Example
```vba
' Create an instance of clsModernStyle class
Set style = New clsModernStyle

' Initialize the styling for your UserForm
style.Initialize Me ' where Me is the UserForm

' The class automatically applies modern styling to all compatible controls on the form
```

## Main Functions

- **Styling Initialization**: The `Initialize` method applies modern styling to all compatible controls on the form
- **Color Configuration**: Ability to configure colors for various controls and states
- **Icon Support**: Using icons from the Segoe MDL2 Assets font for various controls
- **Animations**: Visual feedback during interaction with controls
- **Clear Buttons**: Automatic clear buttons for textboxes and combo boxes

## Working with Controls

The `clsModernStyle` class supports styling of the following controls:
- TextBox
- ComboBox
- ListBox
- CheckBox
- OptionButton
- Frame
- Label
- CommandButton

For each control type, appropriate styling is implemented, taking into account the specific interaction features with the user.

## Style Configuration

The class allows customization of:
- Colors for various controls
- Fonts and text sizes
- Visibility and enabled state of controls
- Interaction behavior (animations, effects)

For color configuration, you can use initialization parameters:
```vba
' Initialize with custom colors
Set style = New clsModernStyle
style.Initialize Me, _
    ColorBarTitleOn:=RGB(0, 100, 200), _
    ColorBarTitleOff:=RGB(120, 120, 120), _
    ColorBarBottomOn:=RGB(0, 100, 200), _
    ColorBarBottomOff:=RGB(180, 180, 180), _
    ColorBackGroundOn:=RGB(255, 255, 255), _
    ColorBackGroundOff:=RGB(245, 245, 245)
```

## Troubleshooting

### Display Issues
- Ensure Microsoft Forms 2.0 Object Library is enabled in references
- Check that controls are added before calling the Initialize method
- Ensure the MultiUse property is set to True for the class

### Interaction Issues
- Check that control events are not overloaded with other handlers
- Ensure control properties are not changed manually while the class is running
- Verify that the class is not initialized multiple times

## License

This project is licensed under the Apache License 2.0 - see the [LICENSE](LICENSE) file for details.