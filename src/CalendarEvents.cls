VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ipm-modules.Calendar.src"
Option Explicit

Public WithEvents DayLabel As MSForms.Label
Attribute DayLabel.VB_VarHelpID = -1
Public WithEvents Arrow As MSForms.Label
Attribute Arrow.VB_VarHelpID = -1

Private Type TCalendarClickEvents
    FirstDayOfMonth As Date
    Form As Object
    Calendar As Calendar
End Type

Private this As TCalendarClickEvents

''' <summary>Форма, на которой размещены элементы управления.</summary>
Public Property Get Form() As Object
    Set Form = this.Form
End Property
Public Property Set Form(ByVal RHS As Object)
    Set this.Form = RHS
End Property

''' <summary>Инстанс класса <c>Calendar</c>.</summary>
Public Property Get Calendar() As Calendar
    Set Calendar = this.Calendar
End Property
Public Property Set Calendar(ByVal RHS As Calendar)
    Set this.Calendar = RHS
End Property

''' <summary>Действие при клике на стрелки.</summary>
Private Sub Arrow_Click()
    this.Calendar.SelectedMonth = IIf(Arrow.Tag = "ArrowUp", this.Calendar.SelectedMonth - 1, this.Calendar.SelectedMonth + 1)
    this.Calendar.DrawForm
End Sub

''' <summary>Действие при клике на число месяца.</summary>
Private Sub DayLabel_Click()
    If Month(DayLabel.Tag) <> this.Calendar.SelectedMonth Then GoTo RepaintForm

    Dim Day As MSForms.Control
    For Each Day In this.Form.Days.Controls
        If Day.BorderStyle = fmBorderStyleSingle Then Day.BorderStyle = fmBorderStyleNone
    Next

    DayLabel.BorderColor = IIf(CStr(DayLabel.Tag) = Date, vbBlack, this.Calendar.ActiveColor)
    DayLabel.BorderStyle = fmBorderStyleSingle

Exit Sub
RepaintForm:
    this.Calendar.SelectedMonth = Month(DayLabel.Tag)
    this.Calendar.DrawForm
End Sub

''' <summary>Действие при двойном клике на число месяца.</summary>
Private Sub DayLabel_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    this.Calendar.SelectedValue = DayLabel.Tag
    this.Form.Hide
End Sub
