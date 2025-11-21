# Техническая документация для проекта VBA Modern Style UserForms

## Содержание
1. [Обзор класса](#обзор-класса)
2. [Архитектура класса](#архитектура-класса)
3. [Свойства](#свойства)
4. [Методы](#методы)
5. [События](#события)
6. [Константы и перечисления](#константы-и-перечисления)
7. [Детали реализации](#детали-реализации)
8. [Зависимости](#зависимости)

## Обзор класса

Проект VBA Modern Style UserForms представляет собой библиотеку классов VBA, предназначенную для стилизации элементов управления MSForms в Excel. Основной класс `clsModernStyle` реализует современные визуальные эффекты, анимации и улучшенную визуальную обратную связь для пользовательских форм.

### Назначение
Класс `clsModernStyle` разработан для применения современного дизайна к элементам управления MSForms с реализациями визуальных эффектов, таких как анимация фокуса, настройка цвета и шрифта, добавление иконок и визуальных элементов.

### Основные возможности
- Применение современного стиля к различным элементам управления (TextBox, ComboBox, ListBox, CheckBox, OptionButton и др.)
- Поддержка анимации фокуса
- Настройка цвета и шрифта
- Добавление иконок и визуальных элементов
- Управление видимостью и состоянием элементов управления

## Архитектура класса

### Обработчики событий
Класс использует события для отслеживания изменений в элементах управления:

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

### Структура данных
Класс хранит информацию об элементах управления в коллекции:
- `mStyleItems` - коллекция всех стилевых элементов
- `mControl` - основной элемент управления
- `mControlType` - тип элемента управления
- `mControlName` - имя элемента управления
- `mControlTipText` - текст всплывающей подсказки

### Основные свойства класса

| Свойство | Тип | Описание |
|----------|-----|----------|
| `control` | MSForms.control | Основной элемент управления, который стилизуется |
| `ControlType` | String | Тип элемента управления (например, "TextBox", "ComboBox") |
| `Name` | String | Имя элемента управления |
| `ControlTipText` | String | Текст всплывающей подсказки для элемента управления |
| `Visible` | Boolean | Видимость элемента управления и всех связанных с ним стилевых элементов |
| `Locked` | Boolean | Состояние блокировки элемента управления и всех связанных с ним стилевых элементов |
| `Enabled` | Boolean | Состояние доступности элемента управления и всех связанных с ним стилевых элементов |
| `top` | Single | Верхняя позиция элемента управления |
| `left` | Single | Левая позиция элемента управления |
| `height` | Single | Высота элемента управления |
| `width` | Single | Ширина элемента управления |
| `FontSizeTitleOff` | Integer | Размер шрифта для неактивного состояния |
| `FontSizeTitleOn` | Integer | Размер шрифта для активного состояния |
| `FontName` | String | Имя шрифта для элемента управления |

## Свойства

### Свойства основного элемента управления
- `control` - Получает или устанавливает основной элемент управления, который будет стилизован
- `ControlType` - Получает или устанавливает тип элемента управления (например, "TextBox", "ComboBox")
- `Name` - Получает или устанавливает имя элемента управления
- `ControlTipText` - Получает или устанавливает текст всплывающей подсказки для элемента управления

### Свойства видимости и состояния
- `Visible` - Получает или устанавливает видимость элемента управления и всех связанных с ним стилевых элементов
- `Locked` - Получает или устанавливает состояние блокировки элемента управления и всех связанных с ним стилевых элементов
- `Enabled` - Получает или устанавливает состояние доступности элемента управления и всех связанных с ним стилевых элементов

### Свойства позиционирования и размеров
- `top` - Получает или устанавливает верхнюю позицию элемента управления
- `left` - Получает или устанавливает левую позицию элемента управления
- `height` - Получает или устанавливает высоту элемента управления
- `width` - Получает или устанавливает ширину элемента управления

### Свойства шрифта
- `FontSizeTitleOff` - Получает или устанавливает размер шрифта для неактивного состояния
- `FontSizeTitleOn` - Получает или устанавливает размер шрифта для активного состояния
- `FontName` - Получает или устанавливает имя шрифта для элемента управления

### Свойства цвета
- `ColorBarTitleOn` - Получает или устанавливает цвет заголовка при активном состоянии
- `ColorBarTitleOff` - Получает или устанавливает цвет заголовка при неактивном состоянии
- `ColorBarBottomOn` - Получает или устанавливает цвет нижней линии при активном состоянии
- `ColorBarBottomOff` - Получает или устанавливает цвет нижней линии при неактивном состоянии
- `ColorBackGroundOn` - Получает или устанавливает цвет фона при активном состоянии
- `ColorBackGroundOff` - Получает или устанавливает цвет фона при неактивном состоянии
- `ColorBarIconOn` - Получает или устанавливает цвет иконки при активном состоянии
- `ColorBarIconOff` - Получает или устанавливает цвет иконки при неактивном состоянии
- `ColorDropArrowOn` - Получает или устанавливает цвет стрелки выпадающего списка при активном состоянии
- `ColorDropArrowOff` - Получает или устанавливает цвет стрелки выпадающего списка при неактивном состоянии
- `ColorTgBorderOn` - Получает или устанавливает цвет границы переключателя при активном состоянии
- `ColorTgBorderOff` - Получает или устанавливает цвет границы переключателя при неактивном состоянии
- `ColorChkBoxBtnOn` - Получает или устанавливает цвет кнопки флажка при активном состоянии
- `ColorChkBoxBtnOff` - Получает или устанавливает цвет кнопки флажка при неактивном состоянии
- `ColorChkBoxCaptionOn` - Получает или устанавливает цвет заголовка флажка при активном состоянии
- `ColorChkBoxCaptionOff` - Получает или устанавливает цвет заголовка флажка при неактивном состоянии

### Свойства символов
- `ChrDropArrowOn` - Получает или устанавливает символ стрелки выпадающего списка при активном состоянии
- `ChrDropArrowOff` - Получает или устанавливает символ стрелки выпадающего списка при неактивном состоянии
- `ChrChkBoxBtnOn` - Получает или устанавливает символ кнопки флажка при активном состоянии
- `ChrChkBoxBtnOff` - Получает или устанавливает символ кнопки флажка при неактивном состоянии
- `ChrOptBoxBtnOn` - Получает или устанавливает символ переключателя при активном состоянии
- `ChrOptBoxBtnOff` - Получает или устанавливает символ переключателя при неактивном состоянии

### Свойства дополнительных элементов
- `BarBottom` - Получает или устанавливает нижнюю линию элемента управления
- `BarTitle` - Получает или устанавливает заголовок элемента управления
- `BarIcon` - Получает или устанавливает иконку элемента управления
- `BackGround` - Получает или устанавливает фон элемента управления
- `DropArrow` - Получает или устанавливает стрелку выпадающего списка
- `BtnClear` - Получает или устанавливает кнопку очистки
- `TgBorder` - Получает или устанавливает границу переключателя
- `ChkBoxBtn` - Получает или устанавливает кнопку флажка
- `ChkBoxCaption` - Получает или устанавливает заголовок флажка

### Свойства коллекции
- `StyleItems` - Получает или устанавливает коллекцию всех стилевых элементов
- `Count` - Получает количество элементов в коллекции
- `getItemByIndex` - Получает элемент из коллекции по индексу
- `getItemByName` - Получает элемент из коллекции по имени
- `Version` - Получает информацию о версии класса

## Методы

### Основные методы инициализации
- `Initialize` - Инициализация стиля для всех элементов управления формы

**Синтаксис:**
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

**Параметры:**
- `Form` - ссылка на UserForm, к которому применяется стиль
- `ColorBarTitleOn` - цвет заголовка в активном состоянии (по умолчанию 14854934)
- `ColorBarTitleOff` - цвет заголовка в неактивном состоянии (по умолчанию 10395294)
- `ColorBarBottomOn` - цвет нижней линии в активном состоянии (по умолчанию 14854934)
- `ColorBarBottomOff` - цвет нижней линии в неактивном состоянии (по умолчанию 10395294)
- `ColorBackGroundOn` - цвет фона в активном состоянии (по умолчанию vbWhite)
- `ColorBackGroundOff` - цвет фона в неактивном состоянии (по умолчанию 1647476)
- `ColorBarIconOn` - цвет иконки в активном состоянии (по умолчанию 14854934)
- `ColorBarIconOff` - цвет иконки в неактивном состоянии (по умолчанию 10395294)
- `ColorDropArrowOn` - цвет стрелки выпадающего списка в активном состоянии (по умолчанию vbBlack)
- `ColorDropArrowOff` - цвет стрелки выпадающего списка в неактивном состоянии (по умолчанию 10395294)
- `ColorTgBorderOn` - цвет границы переключателя в активном состоянии (по умолчанию 14854934)
- `ColorTgBorderOff` - цвет границы переключателя в неактивном состоянии (по умолчанию 10395294)
- `ColorChkBoxBtnOn` - цвет кнопки флажка в активном состоянии (по умолчанию vbBlack)
- `ColorChkBoxBtnOff` - цвет кнопки флажка в неактивном состоянии (по умолчанию 10395294)
- `ChrDropArrowOn` - символ стрелки выпадающего списка в активном состоянии (по умолчанию ArrowOn)
- `ChrDropArrowOff` - символ стрелки выпадающего списка в неактивном состоянии (по умолчанию ArrowOff)
- `ColorChkBoxCaptionOn` - цвет заголовка флажка в активном состоянии (по умолчанию 14854934)
- `ColorChkBoxCaptionOff` - цвет заголовка флажка в неактивном состоянии (по умолчанию 10395294)
- `ChrChkBoxBtnOn` - символ кнопки флажка в активном состоянии (по умолчанию CheckboxComposite)
- `ChrChkBoxBtnOff` - символ кнопки флажка в неактивном состоянии (по умолчанию CheckBox1)
- `ChrOptBoxBtnOn` - символ переключателя в активном состоянии (по умолчанию CircleFill)
- `ChrOptBoxBtnOff` - символ переключателя в неактивном состоянии (по умолчанию CircleRing)

### Методы стилизации элементов управления
- `ApplyControlStyle` - применение стиля в зависимости от типа элемента управления
- `SetCommonStyleProperties` - установка общих свойств стиля для элемента управления
- `setTextBoxStyle` - установка стиля для текстового поля
- `setComboBoxStyle` - установка стиля для комбинированного поля
- `setListBoxStyle` - установка стиля для списка

### Методы добавления стилевых элементов
- `CreateStyledLabel` - создание и настройка основных свойств дополнительного элемента
- `SetCommonFontProperties` - установка общих свойств шрифта для элемента управления
- `addBarBottom` - добавление нижней линии стиля для элемента управления
- `addBarTitle` - добавление заголовка стиля для элемента управления
- `addBarIcon` - добавление иконки стиля для элемента управления
- `addBackGround` - добавление фона стиля для элемента управления
- `addDropArrow` - добавление стрелки выпадающего списка стиля для элемента управления
- `addBtnClear` - добавление кнопки очистки стиля для элемента управления
- `addCheckBox` - добавление стиля флажка для элемента управления
- `addCheckBoxSwitch` - добавление стиля переключателя для элемента управления

### Методы обработки событий
- `HandleExitEvent` - сброс стиля для всех элементов управления
- `exitControl` - сброс стиля элемента управления при потере фокуса
- `btnClearVisible` - управление видимостью кнопки очистки для элемента управления
- `HandleEnterEvent` - активация стиля элемента управления при получении фокуса

## События

### События текстовых полей
- `mTextBox_Change` - событие изменения текста в текстовом поле
- `mTextbox_MouseDown` - событие нажатия мыши на текстовом поле
- `mTextbox_KeyUp` - событие отпускания клавиши при фокусе на текстовом поле

### События комбинированных полей
- `mComboBox_Change` - событие изменения значения в комбинированном поле
- `mComboBox_KeyUp` - событие отпускания клавиши при фокусе на комбинированном поле
- `mComboBox_MouseDown` - событие нажатия мыши на комбинированном поле

### События списков
- `mListBox_Change` - событие изменения значения в списке
- `mListBox_MouseDown` - событие нажатия мыши на списке
- `mListBox_KeyUp` - событие отпускания клавиши при фокусе на списке

### События других элементов управления
- `mUserForm_Click` - событие клика по пользовательской форме
- `mFrame_Click` - событие клика по фрейму
- `mLabel_Click` - событие клика по метке
- `mCommandButton_Click` - событие клика по командной кнопке

### События специфических элементов
- `mDropArrow_Click` - событие клика по стрелке выпадающего списка
- `mBtnClear_Click` - событие клика по кнопке очистки
- `mChkBoxBtn_Click` - событие клика по кнопке флажка
- `mTgBorder_Click` - событие клика по границе переключателя
- `mCheckBox_Change` - событие изменения состояния флажка
- `mChkBoxCaption_Click` - событие клика по заголовку флажка
- `mOptionButton_Change` - событие изменения состояния переключателя

## Константы и перечисления

### Перечисление иконок
```vba
Public Enum enumIcons
    ArrowOff = &HE011                       ' Стрелка выпадающего списка (выкл)
    ArrowOn = &HE010                        ' Стрелка выпадающего списка (вкл)
    CheckBox1 = 59193                       ' Квадрат (обычный)
    Checkbox14 = 61803                      ' Квадрат (маленький)
    CheckboxComposite = 59194               ' Квадрат с галочкой
    CheckboxComposite14 = 61804             ' Квадрат с галочкой (маленький)
    CheckboxCompositeReversed = 59197       ' Квадрат с галочкой (обратный)
    CheckboxIndeterminateCombo = 61806      ' Квадрат с тире
    CheckboxIndeterminateCombo14 = 61805    ' Квадрат с тире (маленький)
    CheckboxFill = 59195                    ' Квадрат (заполненный)
    CheckMark = 59198                       ' Галочка
    CircleFill = 59963                      ' Круг (заполненный)
    CircleRing = 59962                      ' Круг (контур)
    FavoriteStar = 59188                    ' Звезда (обычная)
    FavoriteStarFill = 59189                ' Звезда (заполненная)
    Heart = 60241                           ' Сердце (обычное)
    HeartFill = 60242                       ' Сердце (заполненное)
    InkingColorFill = 60775                 ' Кисть (заполненная)
    InkingColorOutline = 60774              ' Кисть (контур)
    PaginationDotOutline10 = 61734          ' Точка (контур)
    PaginationDotSolid10 = 61735            ' Точка (заполненная)
    PasswordChar = 149                      ' Символ для скрытия пароля
    RadioBtnOff = 60618                     ' Радио кнопка (выкл)
    RadioBtnOn = 60619                      ' Радио кнопка (вкл)
    ToggleOff = 60434                       ' Переключатель (выкл)
    ToggleOn = 60433                        ' Переключатель (вкл)
    ToggleThumb = 60436                     ' Ползунок переключателя
End Enum
```

### Константы
```vba
' Константы шрифта
Private Const FONT_NAME_ICON As String = "Segoe MDL2 Assets"

' Константы для типов элементов управления
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

' Константы для дополнительных имен элементов управления
Private Const BAR_BOTTOM As String = "_barBottom"
Private Const BAR_TITLE As String = "_barTitle"
Private Const BAR_ICON As String = "_barIcon"
Private Const BACK_GROUND As String = "_BackGround"
Private Const DROP_ARROW As String = "_DropArrow"
Private Const BTN_CLEAR As String = "_BtnClear"

' Константы для поведения элементов управления
Private Const CONTROL_SWITCH As String = "SWITCH"
```

## Детали реализации

### Вспомогательные методы
- `UpdateSwitchState` - внутренний метод для обновления состояния переключателя
- `UpdateSwitchVisualState` - внутренний метод для обновления визуального состояния переключателя
- `IsControlActive` - вспомогательный метод для проверки активности элемента управления
- `ConfigureStyleElement` - внутренний метод для настройки свойств элемента стиля
- `IsControlInCollection` - проверка наличия элемента управления в коллекции
- `SetControlEnabled` - внутренний метод для установки состояния доступности элемента управления
- `SetControlVisibility` - внутренний метод для установки видимости элемента управления
- `SetControlLock` - внутренний метод для установки состояния блокировки элемента управления

### Метод очистки
- `Class_Terminate` - очистка объектов при завершении работы класса

## Зависимости

- Microsoft Forms 2.0 Object Library
- Среда выполнения VBA