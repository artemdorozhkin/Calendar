## Calendar

Calendar - это небольшая библиотека на Visual Basic for Applications. Объект календаря может быть использован на формах пользователя для отображения и взаимодействия с датами.

![Calendar](<assets/Screenshot 2023-06-29 081231.png>)

### Чтобы начать: функция NewCalendar

Функция NewCalendar создает и инициализирует экземпляр класса Calendar. Она принимает следующие параметры:

```vb
Public Function NewCalendar(ByRef Form As Object, Optional ByVal Caption As String = "Календарь", Optional ByVal FontSize As Integer, Optional ByVal SelectedValue As String, Optional ByVal SelectedMonth As Integer = -1, Optional ByVal LabelSize As Integer = 18, Optional ByVal ActiveColor As Long = &HC000&) As Calendar
```

- `Form`: Ссылка на объект формы, на которой будет размещен календарь.
- `Caption` (Необязательный): Строка, определяющая заголовок или название календаря. Значение по умолчанию - "Календарь".
- `FontSize` (Необязательный): Целое число, определяющее размер шрифта для меток дней календаря.
- `SelectedValue` (Необязательный): Строка, представляющая выбранное значение в календаре. Если не указано, используется текущая дата.
- `SelectedMonth` (Необязательный): Целое число, представляющее выбранный месяц в календаре. Если не указано, используется месяц, соответствующий выбранному значению `SelectedValue`.
- `LabelSize` (Необязательный): Целое число, определяющее размер меток дней календаря. Значение по умолчанию - 18.
- `ActiveColor` (Необязательный): Целое число (`Long`), представляющее цвет активного дня в календаре. Значение по умолчанию - `&HC000&` (светло-зеленый).

#### Пример использования

Вот пример использования функции NewCalendar в вашем проекте на VBA:

```vb
Private Sub UserForm_Initialize()
    Dim Calendar As Calendar
    Set Calendar = NewCalendar(myForm, "Мой календарь", 12, "2023-06-30", 6, 20)
End Sub
```

В приведенном примере мы создаем экземпляр класса Calendar с помощью функции NewCalendar и передаем ссылку на форму, заголовок, размер шрифта, выбранное значение, выбранный месяц, размер меток.

### Заметки

- Для использования этого модуля необходимо добавить его в ваш проект VBA.
- Вы можете настроить параметры функции NewCalendar в соответствии с вашими потребностями и требованиями проекта.
- Вы также можете расширить функциональность класса Calendar, добавив новые методы и свойства по своему усмотрению.