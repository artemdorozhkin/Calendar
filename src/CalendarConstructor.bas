Attribute VB_Name = "CalendarConstructor"
'@Folder "ipm-modules.Calendar.src"
Option Explicit

Public Function NewCalendar(ByRef Form As Object, Optional ByVal Caption As String = "Календарь", Optional ByVal FontSize As Integer, Optional ByVal SelectedValue As String, Optional ByVal SelectedMonth As Integer = -1, Optional ByVal LabelSize As Integer = 18, Optional ByVal ActiveColor As Long = &HC000&) As Calendar
    Set NewCalendar = New Calendar

    With NewCalendar
        Set .Form = Form
        .Caption = Caption
        .FontSize = FontSize
        .SelectedValue = IIf(Len(SelectedValue) = 0, Date, SelectedValue)
        .SelectedMonth = IIf(SelectedMonth = -1, DateTime.Month(.SelectedValue), SelectedMonth)
        .Size = LabelSize
        .ActiveColor = ActiveColor

        .DrawForm
    End With
End Function
