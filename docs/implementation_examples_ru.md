# Примеры реализации для проекта VBA Modern Style UserForms

## Содержание
1. [Базовая реализация](#базовая-реализация)
2. [Расширенная настройка цветов](#расширенная-настройка-цветов)
3. [Создание переключателей](#создание-переключателей)
4. [Добавление иконок](#добавление-иконок)
5. [Динамическая настройка стилей](#динамическая-настройка-стилей)
6. [Работа с коллекцией стилей](#работа-с-коллекцией-стилей)
7. [Интеграция с существующими формами](#интеграция-с-существующими-формами)
8. [Примеры обработки событий](#примеры-обработки-событий)
9. [Создание тем оформления](#создание-тем-оформления)
10. [Продвинутые примеры](#продвинутые-примеры)

## Базовая реализация

### Простой пример использования
```vba
' В модуле формы
Dim MStyleItem As clsModernStyle

Private Sub UserForm_Initialize()
    ' Создание экземпляра класса стиля
    Set MStyleItem = New clsModernStyle
    
    ' Инициализация стилей для текущей формы
    Call MStyleItem.Initialize(Me)
End Sub
```

### Пример с несколькими элементами управления
```vba
Private Sub UserForm_Initialize()
    ' Добавление элементов управления на форму через конструктор
    ' (TextBox1, ComboBox1, CheckBox1, OptionButton1)
    
    Set MStyleItem = New clsModernStyle
    
    ' Установка всплывающих подсказок для элементов управления
    TextBox1.ControlTipText = "Имя пользователя"
    ComboBox1.ControlTipText = "Выберите опцию"
    
    ' Инициализация стилей
    Call MStyleItem.Initialize(Me)
End Sub
```

## Расширенная настройка цветов

### Настройка цветов при инициализации
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Инициализация с настраиваемыми цветами
    MStyleItem.Initialize Me, _
        ColorBarTitleOn:=RGB(0, 100, 200), _      ' Цвет активного заголовка
        ColorBarTitleOff:=RGB(120, 120, 120), _    ' Цвет неактивного заголовка
        ColorBarBottomOn:=RGB(0, 100, 200), _      ' Цвет активной нижней линии
        ColorBarBottomOff:=RGB(180, 180, 180), _   ' Цвет неактивной нижней линии
        ColorBackGroundOn:=RGB(255, 255, 255), _   ' Цвет активного фона
        ColorBackGroundOff:=RGB(245, 245, 245), _  ' Цвет неактивного фона
        ColorBarIconOn:=RGB(0, 100, 200), _        ' Цвет активной иконки
        ColorBarIconOff:=RGB(150, 150, 150), _     ' Цвет неактивной иконки
        ColorDropArrowOn:=RGB(0, 100, 200), _      ' Цвет активной стрелки
        ColorDropArrowOff:=RGB(150, 150, 150)      ' Цвет неактивной стрелки
End Sub
```

### Использование предопределенных цветовых схем
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Цветовая схема "Темная тема"
    If ThemeManager.IsDarkTheme Then
        MStyleItem.Initialize Me, _
            ColorBarTitleOn:=RGB(255, 255, 255), _
            ColorBarTitleOff:=RGB(180, 180, 180), _
            ColorBarBottomOn:=RGB(0, 120, 215), _
            ColorBarBottomOff:=RGB(80, 80, 80), _
            ColorBackGroundOn:=RGB(30, 30, 30), _
            ColorBackGroundOff:=RGB(20, 20, 20)
    Else
        ' Цветовая схема "Светлая тема"
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

## Создание переключателей

### Простой переключатель
```vba
Private Sub UserForm_Initialize()
    ' Установка свойства Tag для создания переключателя
    CheckBox1.Tag = "SWITCH"
    CheckBox2.Tag = "SWITCH"
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Группа переключателей
```vba
Private Sub UserForm_Initialize()
    ' Создание группы переключателей
    ToggleOption1.Caption = "Опция 1"
    ToggleOption1.Tag = "SWITCH"
    
    ToggleOption2.Caption = "Опция 2"
    ToggleOption2.Tag = "SWITCH"
    
    ToggleOption3.Caption = "Опция 3"
    ToggleOption3.Tag = "SWITCH"
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Установка начального состояния
    ToggleOption1.Value = True
End Sub
```

### Динамическое создание переключателей
```vba
Private Sub UserForm_Initialize()
    ' Динамическое создание переключателей
    Dim chk As MSForms.CheckBox
    Dim i As Integer
    
    For i = 1 To 5
        Set chk = Me.Controls.Add("Forms.CheckBox.1", "DynamicToggle" & i, True)
        With chk
            .Left = 20
            .Top = 50 + (i - 1) * 30
            .Width = 200
            .Height = 20
            .Caption = "Переключатель " & i
            .Tag = "SWITCH"
        End With
    Next i
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

## Добавление иконок

### Использование встроенных иконок
```vba
Private Sub UserForm_Initialize()
    ' Установка иконок через числовые значения из enumIcons
    TextBox1.Tag = 59193  ' CheckBox1
    TextBox2.Tag = 59188  ' FavoriteStar
    TextBox3.Tag = 60241  ' Heart
    
    ComboBox1.Tag = 61735  ' PaginationDotSolid10
    ListBox1.Tag = 59962   ' CircleRing
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Использование иконок для разных типов элементов управления
```vba
Private Sub UserForm_Initialize()
    ' Иконки для текстовых полей
    UsernameBox.Tag = 59193   ' Квадрат для имени пользователя
    PasswordBox.Tag = 149     ' Символ пароля
    EmailBox.Tag = 59188      ' Звезда для email
    
    ' Иконки для комбинированных полей
    CountryCombo.Tag = 60619  ' Радио кнопка для страны
    CategoryCombo.Tag = 61804 ' Квадрат с галочкой для категории
    
    ' Иконки для флажков
    AgreementBox.Tag = 59194  ' Квадрат с галочкой для соглашения
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

### Динамическое добавление иконок
```vba
Private Sub UserForm_Initialize()
    Dim controlsList As Variant
    Dim iconsList As Variant
    Dim i As Integer
    
    ' Список элементов управления и соответствующих иконок
    controlsList = Array("TextBox1", "TextBox2", "ComboBox1", "CheckBox1")
    iconsList = Array(59193, 59188, 60619, 59194) ' Значения из enumIcons
    
    ' Применение иконок
    For i = 0 To UBound(controlsList)
        Me.Controls(controlsList(i)).Tag = iconsList(i)
    Next i
    
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub
```

## Динамическая настройка стилей

### Изменение стилей после инициализации
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Изменение стилей конкретных элементов после инициализации
    With MStyleItem.getItemByName(TextBox1.Name)
        .ColorBackGroundOff = RGB(255, 250, 200) ' Желтый фон
        .ColorBackGroundOn = RGB(255, 25, 255)   ' Белый фон при фокусе
        .ColorBarBottomOn = RGB(255, 0, 0)        ' Красная линия при фокусе
    End With
End Sub

Private Sub ChangeStyleButton_Click()
    ' Динамическое изменение стиля по нажатию кнопки
    With MStyleItem.getItemByName(TextBox2.Name)
        .ColorBarTitleOn = RGB(Rnd * 255, Rnd * 25, Rnd * 255)
        .ColorBarBottomOn = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    End With
End Sub
```

### Условная настройка стилей
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Условная настройка стилей в зависимости от типа данных
    Dim ctrl As MSForms.control
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "TextBox" Then
            With MStyleItem.getItemByName(ctrl.Name)
                ' Если имя элемента содержит "password", установить специальный стиль
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

## Работа с коллекцией стилей

### Перебор всех элементов стиля
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Перебор всех элементов стиля и вывод информации
    Dim item As clsModernStyle
    For Each item In MStyleItem.StyleItems
        Debug.Print "Элемент: " & item.Name & ", Тип: " & item.ControlType
    Next item
End Sub
```

### Поиск элементов по критериям
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
    
    ' Найти все текстовые поля
    Dim textBoxes As Collection
    Set textBoxes = FindControlsByType("TextBox")
    
    ' Применить специальный стиль ко всем текстовым полям
    Dim item As clsModernStyle
    For Each item In textBoxes
        With item
            .ColorBarBottomOn = RGB(0, 100, 200)
            .ColorBarBottomOff = RGB(150, 150, 150)
        End With
    Next item
End Sub
```

### Групповая настройка элементов
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Список имен элементов для групповой настройки
    Dim importantFields As Variant
    importantFields = Array("UsernameBox", "PasswordBox", "EmailBox")
    
    ' Применить красную рамку к важным полям
    Dim i As Integer
    For i = 0 To UBound(importantFields)
        On Error Resume Next
        With MStyleItem.getItemByName(importantFields(i))
            .ColorBarBottomOn = RGB(25, 0, 0)
            .ColorBarBottomOff = RGB(200, 0, 0)
        End With
        On Error GoTo 0
    Next i
End Sub
```

## Интеграция существующими формами

### Добавление стилей к существующей форме
```vba
' В отдельном модуле
Public Sub ApplyModernStyleToForm(formName As String)
    Dim frm As Object
    Set frm = VBA.Interaction.CreateObject("Forms." & formName)
    
    Dim style As New clsModernStyle
    style.Initialize frm
    
    frm.Show
End Sub

' Использование
Private Sub CommandButton1_Click()
    ApplyModernStyleToForm "ExistingForm"
End Sub
```

### Постепенная стилизация формы
```vba
Private Sub UserForm_Initialize()
    ' Инициализация без стилизации
    Set MStyleItem = New clsModernStyle
    
    ' Добавление элементов управления программно
    AddStyledControl "TextBox", "UserInput", 50, 50, 200, 20
    AddStyledControl "ComboBox", "SelectionBox", 50, 80, 200, 20
    AddStyledControl "CheckBox", "AgreementBox", 50, 110, 200, 20
End Sub

Private Sub AddStyledControl(controlType As String, controlName As String, _
                           leftPos As Single, topPos As Single, _
                           widthSize As Single, heightSize As Single)
    Dim newControl As MSForms.control
    
    ' Создание элемента управления
    Set newControl = Me.Controls.Add("Forms." & controlType & ".1", controlName, True)
    With newControl
        .left = leftPos
        .top = topPos
        .width = widthSize
        .height = heightSize
    End With
    
    ' Инициализация стилей только для нового элемента
    ' (требуется модификация класса для поддержки одиночной стилизации)
    Call MStyleItem.Initialize(Me)
End Sub
```

## Примеры обработки событий

### Обработка событий стилизованных элементов
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
End Sub

Private Sub TextBox1_Change()
    ' Обработка изменения текста в стилизованном элементе
    If Len(TextBox1.Value) > 0 Then
        ' Изменение стиля в зависимости от содержимого
        With MStyleItem.getItemByName(TextBox1.Name)
            If IsValidEmail(TextBox1.Value) Then
                .ColorBarBottomOn = RGB(0, 150, 0)   ' Зеленый при валидном email
            Else
                .ColorBarBottomOn = RGB(255, 0, 0)   ' Красный при невалидном
            End If
        End With
    End If
End Sub

Private Function IsValidEmail(email As String) As Boolean
    ' Простая проверка email
    IsValidEmail = (InStr(email, "@") > 0 And InStr(email, ".") > 0)
End Function
```

### События для переключателей
```vba
Private Sub ToggleOption1_Change()
    If ToggleOption1.Value Then
        ' Изменение стиля других элементов при изменении переключателя
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

## Создание тем оформления

### Менеджер тем
```vba
' В отдельном классе или модуле
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
            theme.BackgroundActive = RGB(25, 255, 255)
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
    currentTheme = GetTheme("BlueTheme")  ' Или "GreenTheme", "RedTheme"
    
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

### Переключение тем
```vba
Private Sub ThemeComboBox_Change()
    ApplyTheme ThemeComboBox.Value
End Sub

Private Sub ApplyTheme(themeName As String)
    Dim theme As ThemeColors
    theme = GetTheme(themeName)
    
    ' Применение темы ко всем элементам
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

## Продвинутые примеры

### Форма входа с валидацией
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    Call MStyleItem.Initialize(Me)
    
    ' Установка всплывающих подсказок
    UsernameBox.ControlTipText = "Имя пользователя"
    PasswordBox.ControlTipText = "Пароль"
    LoginButton.Caption = "Войти"
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
    ' Проверка валидности всех полей перед отправкой
    If IsValidForm Then
        MsgBox "Вход выполнен успешно!", vbInformation
        Unload Me
    Else
        MsgBox "Пожалуйста, проверьте правильность заполнения полей.", vbExclamation
    End If
End Sub

Private Function IsValidForm() As Boolean
    IsValidForm = (Len(UsernameBox.Value) >= 3 And Len(PasswordBox.Value) >= 6)
End Function
```

### Динамическая форма с добавлением полей
```vba
Private Sub UserForm_Initialize()
    Set MStyleItem = New clsModernStyle
    
    ' Инициализация с базовыми элементами
    Call MStyleItem.Initialize(Me)
    
    ' Добавление динамических полей
    AddDynamicField "Имя", "firstName", 20, 100
    AddDynamicField "Фамилия", "lastName", 20, 140
    AddDynamicField "Email", "email", 20, 180
End Sub

Private Sub AddDynamicField(promptText As String, fieldName As String, _
                          leftPos As Single, topPos As Single)
    ' Создание текстового поля
    Dim newTextBox As MSForms.TextBox
    Set newTextBox = Me.Controls.Add("Forms.TextBox.1", fieldName, True)
    
    With newTextBox
        .left = leftPos
        .top = topPos
        .width = 200
        .height = 20
        .ControlTipText = promptText
    End With
    
    ' Повторная инициализация стилей для добавленных элементов
    Call MStyleItem.Initialize(Me)
End Sub

Private Sub AddFieldButton_Click()
    ' Добавление нового поля по нажатию кнопки
    Static fieldCounter As Integer
    fieldCounter = fieldCounter + 1
    
    AddDynamicField "Дополнительное поле " & fieldCounter, _
                    "extraField" & fieldCounter, 20, 180 + fieldCounter * 40
End Sub
```

Эти примеры показывают различные способы использования класса `clsModernStyle` для создания современных и функциональных пользовательских форм в Excel. Каждый пример можно адаптировать под конкретные потребности приложения.