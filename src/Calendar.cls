VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ipm-modules.Calendar.src"
Option Explicit

Public Enum EDraw
    Days
    Months
End Enum

Private Type TCalendar
    Controls As Collection
    Form As Object
    Days As MSForms.Frame
    CurrentMY As MSForms.Label
    CalendarEvents As CalendarEvents
    Size As Integer
    Month As Integer
    Value As Date
    MonthDay As Date
    Name As String
    MonthNames As Object
    FontSize As Integer
    ActiveColor As Long
End Type

Private this As TCalendar

Private Sub Class_Initialize()
    this.ActiveColor = &HC000&
End Sub

''' <summary>Коллекция элементов управления, добавленных на форму. Включает элемент класса <c>CalendarEvents</c>.</summary>
Public Property Get Controls() As Collection
    Set Controls = this.Controls
End Property

''' <summary>Пользовательская форма для размещения календаря.</summary>
Public Property Get Form() As Object
    Set Form = this.Form
End Property
Public Property Set Form(ByRef RHS As Object)
    Set this.Form = RHS

    SetFormatForm
End Property

''' <summary>Основной цвет.</summary>
Public Property Get ActiveColor() As Long
    ActiveColor = this.ActiveColor
End Property
Public Property Let ActiveColor(ByVal RHS As Long)
    this.ActiveColor = RHS
End Property

''' <summary>Название формы.</summary>
Public Property Get Caption() As String
    Caption = this.Form.Caption
End Property
Public Property Let Caption(ByVal RHS As String)
    this.Form.Caption = RHS
End Property

''' <summary>Размер шрифта чисел месяца.</summary>
Public Property Get FontSize() As Integer
    FontSize = IIf(this.FontSize = 0, 10, this.FontSize)
End Property
Public Property Let FontSize(ByVal RHS As Integer)
    this.FontSize = RHS
End Property

''' <summary>Текущий выбранный месяц.</summary>
Public Property Get SelectedMonth() As Integer
    SelectedMonth = IIf(this.Month = 0, Month(Date), this.Month)
End Property
Public Property Let SelectedMonth(ByVal RHS As Integer)
    this.Month = RHS
End Property

''' <summary>Текущая выбранная дата.</summary>
Public Property Get SelectedValue() As Date
    SelectedValue = IIf(this.Value = 0, Date, this.Value)
End Property
Public Property Let SelectedValue(ByVal RHS As Date)
    this.Value = RHS
End Property

''' <summary>Размер элементов управления чисел месяца.</summary>
Public Property Get Size() As Integer
    Size = IIf(this.Size = 0, 18, this.Size)
End Property
Public Property Let Size(ByVal RHS As Integer)
    this.Size = RHS
End Property

''' <summary>Процедура для инициализации нового инстанса класса <c>CalendarEvents</c>.</summary>
Private Sub InitNewCalendarEvents()
    Set this.CalendarEvents = New CalendarEvents
    Set this.CalendarEvents.Calendar = Me
    Set this.CalendarEvents.Form = this.Form
End Sub

''' <summary>Устанавливает форматирование формы с добавлением рамки для чисел месяца.</summary>
Private Sub SetFormatForm()
    With this.Form
        .Caption = "Календарь"
        .BackColor = vbHighlightText
        .Height = 260
        .Width = 253
        With .Controls.Add("Forms.Frame.1", "Days")
            .Height = 138
            .Width = 216
            .Left = 12
            .Top = 66
            .SpecialEffect = 0
        End With
    End With
End Sub

''' <summary>Основная процедура для отрисовки формы.</summary>
''' <param name="WhatDraw">Пока не работает. Выбор отрисовки дней, месяцев или лет.</param>
Public Sub DrawForm(Optional WhatDraw As EDraw = EDraw.Days)
    If WhatDraw = Days Then DrawDays
    If WhatDraw = Months Then Exit Sub

    DrawDaysOfWeek
    DrawCurrentMonthYear
    DrawArrows
End Sub

''' <summary>Отрисовка стрелок для перемещения по месяцам.</summary>
Private Sub DrawArrows()
    Dim Arrows As Variant
    Arrows = Array(ChrW(8593), ChrW(8595))

    Dim i As Byte
    For i = LBound(Arrows) To UBound(Arrows)
        Dim Arrow As MSForms.Label
        Set Arrow = Me.Form.Controls.Add("Forms.Label.1", "Arrow" & i)
        With Arrow
            .Tag = "Arrow" & IIf(i = 0, "Up", "Down")
            .Caption = Arrows(i)
            .Font.Size = 14
            .TextAlign = fmTextAlignCenter
            .Left = IIf(i = 0, 180, 204)
            .Top = 18
            .Width = 18
            .Height = 18
        End With
        InitNewCalendarEvents
        Set this.CalendarEvents.Arrow = Arrow:
        this.Controls.Add this.CalendarEvents, Arrow.Name
    Next
End Sub

''' <summary>Отрисовка лейбла с указанием выбранного месяца и года.</summary>
Private Sub DrawCurrentMonthYear()
    InitMonthNames
    Set this.CurrentMY = Me.Form.Controls.Add("Forms.Label.1", "CurrentMonthYear")
    With this.CurrentMY
        .Caption = GetCurrentMonthYear
        .Font.Size = 14
        .Left = 18
        .Top = 18
        .Width = 126
        .Height = 18
    End With
End Sub

''' <summary>Отрисовка дней недели.</summary>
Private Sub DrawDaysOfWeek()
    Dim DaysOfWeek As Variant
    DaysOfWeek = Array("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")
    Dim Left As Integer
    Left = 24

    Dim i As Long
    For i = LBound(DaysOfWeek) To UBound(DaysOfWeek)
        Dim CurrentDay As MSForms.Label
        Set CurrentDay = Me.Form.Controls.Add("Forms.Label.1", "DayOfWeek" & i)
        With CurrentDay
            .Caption = DaysOfWeek(i)
            .Font.Size = 10
            Left = IIf(i = LBound(DaysOfWeek), Left, Left + 30)
            .Left = Left
            .Top = 48
            .Width = 18
            .Height = 18
            .TextAlign = fmTextAlignCenter
        End With
    Next
End Sub

''' <summary>Функция для получения выбранного месяца и года.</summary>
''' <returns>
''' Строку формата "Месяц год".
''' <example>"Сентябрь 2022"</example>
''' </returns>
Private Function GetCurrentMonthYear() As String
    Dim FirstDay As Date
    FirstDay = FirstDayOfMonth
    InitMonthNames

    GetCurrentMonthYear = this.MonthNames(Month(FirstDay)) & Space(1) & Year(FirstDay)
End Function

''' <summary>Инициализирует словарь с названием месяцев, где ключ - номер месяца.</summary>
Private Sub InitMonthNames()
    Dim Names As Variant
    Names = Array("Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь")

    Set this.MonthNames = CreateObject("Scripting.Dictionary")
    Dim i As Byte
    For i = LBound(Names) To UBound(Names)
        this.MonthNames.Add Item:=Names(i), Key:=i + 1
    Next
End Sub

''' <summary>Отрисовка чисел месяца.</summary>
Private Sub DrawDays()
    Dim MonthDays As Variant
    MonthDays = GetMonthDays()
    Set this.Controls = New Collection
    Me.Form.Days.Controls.Clear

    Dim i As Integer
    For i = LBound(MonthDays) To UBound(MonthDays)
        this.MonthDay = MonthDays(i)
        this.Name = "Day" & i
        Dim DayLabel As MSForms.Label
        Set DayLabel = AddLabel()

        InitNewCalendarEvents
        Set this.CalendarEvents.DayLabel = DayLabel
        this.Controls.Add this.CalendarEvents, CStr(this.MonthDay)
    Next
    SetFormatSelected
End Sub

''' <summary>Установка формата выбранному числу месяца.</summary>
Private Sub SetFormatSelected()
    Dim Day As MSForms.Control
    For Each Day In this.Form.Days.Controls
        If Day.BorderStyle = fmBorderStyleSingle Then Day.BorderStyle = fmBorderStyleNone
        If CDate(Day.Tag) <> Me.SelectedValue Then GoTo NextDay
        Day.ForeColor = IIf(Day.Tag = Date, vbWhite, this.ActiveColor)
NextDay:
    Next
End Sub

''' <summary>Возвращает первое число переданного месяца.</summary>
''' <remarks>Если значение <c>Month</c> не передано, за основу берется текущий месяц.</remarks>
''' <param name="Month">Необзяательный. Номер месяца.</param>
Public Function FirstDayOfMonth(Optional Month As Integer) As Date
    FirstDayOfMonth = DateTime.DateSerial(Year(Date), SelectedMonth, 1)
End Function

''' <summary>Получение чисел месяца, начиная с понедельника.</summary>
''' <remarks>
''' Если первый день выбранного месяца пришелся на вторник, массив начнется с последнего дня предыдущего месяца и т.д.
''' Если последний день выбранного месяца пришелся на субботу, массив закончится первым днем следующего месяца и т.д.
''' </remarks>
''' <returns>Возвращает массив с числами месяца.</returns>
Private Function GetMonthDays() As Variant
    Dim FirstMonday As Date
    FirstMonday = FirstDayOfMonth - WorksheetFunction.Weekday(FirstDayOfMonth, 3)

    Dim MonthDays(): ReDim MonthDays(0 To 41)
    Dim i As Integer
    For i = LBound(MonthDays) To UBound(MonthDays)
        If i = LBound(MonthDays) Then MonthDays(i) = FirstMonday: GoTo NextDay
        MonthDays(i) = MonthDays(i - 1) + 1
NextDay:
    Next

    GetMonthDays = MonthDays
End Function

''' <summary>Добавляет лейбл числа месяца.</summary>
''' <returns>Возвращает добавленный лейбл.</returns>
Private Function AddLabel() As MSForms.Label
    Dim DayLabel As MSForms.Label
    Set DayLabel = Me.Form.Days.Controls.Add("Forms.Label.1", this.Name)
    DayLabel.Caption = DateTime.Day(this.MonthDay)
    DayLabel.Tag = this.MonthDay
    FormatLabel DayLabel

    SetPosition DayLabel
    Set AddLabel = DayLabel
End Function

''' <summary>Добавляет формат переданному лейблу.</summary>
''' <param name="DayLabel">Лейбл для форматирования.</param>
Private Sub FormatLabel(ByVal DayLabel As MSForms.Label)
    DayLabel.Width = Me.Size
    DayLabel.Height = Me.Size
    DayLabel.Font.Size = Me.FontSize
    DayLabel.TextAlign = fmTextAlignCenter

    If Month(this.MonthDay) <> Month(FirstDayOfMonth) Then DayLabel.ForeColor = vbScrollBars
    If this.MonthDay = Date Then FormatCurrentDay DayLabel
End Sub

''' <summary>Добавляет формат текущего дня переданному лейблу.</summary>
''' <param name="DayLabel">Лейбл для форматирования.</param>
Private Sub FormatCurrentDay(ByVal DayLabel As MSForms.Label)
    DayLabel.BackColor = &HC000&
    DayLabel.ForeColor = vbWhite
    DayLabel.BorderColor = vbBlack
    DayLabel.BorderStyle = fmBorderStyleSingle
End Sub

''' <summary>Статическая процедура. Устанавливает позицию переданному лейблу.</summary>
''' <param name="DayLabel">Лейбл для установки позиции.</param>
Private Static Sub SetPosition(ByVal DayLabel As MSForms.Label)
    Dim Left As Integer
    If Left = 192 Then Left = 0
    Left = IIf(Left = 0, 12, Left + 30)
    Dim Top As Integer
    If Top = 126 And Left = 12 Then Top = 0
    Top = IIf(Top = 0, 6, IIf(Left = 12, Top + 24, Top))

    DayLabel.Top = Top
    DayLabel.Left = Left
End Sub
