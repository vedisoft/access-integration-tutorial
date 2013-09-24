Пример интеграции MS Access с сервисом Простые Звонки
==========================================================

Простые Звонки - сервис для интеграции клиентских приложений (Excel, 1C и ERP-cистем) с офисными и облачными АТС. Клиентское приложение может общаться с сервером Простых Звонков через единый API, независимо от типа используемой АТС. 

В данном примере мы рассмотрим процесс подключения к серверу Простых Звонков базы данных MS Access.

Мы возмём за основу стандартный шаблон MS Access под названием "Контакты" и добавим в него базовые функции:

- отображение всплывающей карточки при входящем звонке;
- звонок клиенту по клику на телефоный номер;
- умная переадресация на менеджера клиента;

Шаг 0. Установка ActiveX
--------------------------

Необходимо скачать ActiveX по ссылке [отсюда](http://prostiezvonki.ru/installs/ProstieZvonki_ActiveX.exe)

После установки необходимо подключить ActiveX в редакторе VB кода в Tools -> References :

![Включаем ActiveX](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/references.png)

Шаг 1. Настройка подключения к тестовому серверу
--------------------------------------

Теперь нужно скачать [тестовый сервер и диагностическую утилиту](https://github.com/vedisoft/pz-developer-tools).

Запустим тестовый сервер:

    > TestServer.exe -r

и подключимся к нему диагностической утилитой:

    > Diagnostic.exe

    [events off]> Connect ws://localhost:10150 asd
    Успешно начато установление соединения с АТС

Тестовое окружение настроено.

Шаг 2. Создадим VB класс для взаимодействия с ActiveX
------------------------------------------------------

Полный текст модуля, включающий обработку событий, находится в репозитории в файле **ProstieZvonkiWrapper.cls**


Добавим метод для инициализации библиотеки и подключения:

```vb
Option Explicit

Public WithEvents prostie_zvonki_lib As CTIControlX

Public Function Initialize(Phone As String)
    Set prostie_zvonki_lib = New CTIControlX
    On Error GoTo Errhandler
    Call prostie_zvonki_lib.Connect("127.0.0.1:10150", Phone, "Excel", "a9548dc6-4e09-4faa-8cfa-8b5fbbb03087", Environ("LocalAppData") & "\Ведисофт\Excel\ProtocolLib_log.log", 2, 5000)
    Exit Function
Errhandler:
    MsgBox ("Can't connect to server")
End Function

Public Sub Class_Terminate()
    If Not IsNull(prostie_zvonki_lib) Then
        Call prostie_zvonki_lib.Disconnect
    End If
End Sub
```

А также метод для совершения исходящего вызова, который будем использовать при нажатии на кнопку на форме:

```vb
Public Sub MakeCall(Phone As String)
    Call prostie_zvonki_lib.Call("123", Phone)
End Sub
```

Шаг 3. Добавим метод для обработки события входящего звонка
-----------------------------------------------------------

Для примера, будем показывать простое информационное сообщение, которое будет отображать номер и имя звонящего (если найден в базе данных).
Имя будем искать по номеру телефона (колонка Рабочий телефон) в таблице Контакты.

```vb
Private Sub prostie_zvonki_lib_OnTransferredCall(ByVal CallID As Long, ByVal src As String, ByVal dst As String)
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT [Имя], [Фамилия] FROM Контакты WHERE Контакты.[Рабочий телефон] = '" & src & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        MsgBox ("Incoming call from " & rs.Fields(0).Value & " " & rs.Fields(1).Value) 'show client name
    Else
        MsgBox ("Incoming call from " & src) 'can't find client, show number only
    End If
    Call rs.Close
End Sub
```

![Входящий звонок](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/incoming_call.png)


Шаг 4. Добавим метод для обработки события OnTransferredRequest, используемый для умной переадресации
-----------------------------------------------------------------------------------------------------

```vb
Private Sub prostie_zvonki_lib_OnTransferRequest(ByVal CallID As Long, ByVal from As String)
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT Менеджеры.[Телефон] FROM Менеджеры INNER JOIN Контакты ON Менеджеры.[Код] = Контакты.[Менеджер] WHERE Контакты.[Рабочий телефон] = '" & from & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        Call prostie_zvonki_lib.Transfer(CallID, rs.Fields(0).Value) 'call with number, if manager can handle call
    Else
        Call prostie_zvonki_lib.Transfer(CallID, "") 'call with empty string, if manager can't handle call
    End If
    Call rs.Close
End Sub
```

Шаг 5. Создадим VB модуль для доступа к объекту класса и его методам
---------------------------------------------------

Полный текст модуля находится в репозитории в файле **ProstieZvonki.bas**

```vb
Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Function Init()
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
End Function

Public Function MakeCall(Phone As String)
    Call prostie_zvonki_wrapper.MakeCall(Phone)
End Function
```

Шаг 6. Инициализация модуля при открытии базы данных
----------------------------------------------------
Необходимо создать макрос со специальным именем AutoExec.
Этот макрос будет всегда запускаться при открытии базы данных.

![Макрос AutoExec](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/autoexec_macros.png)


Шаг 7. Добавим обработчик для гиперссылки номера телефона в списке клиентов через обработку события Click
-----------------------------------------------------------------------------------------------------------
![Ссылка](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/hiperlink.png)

Полный текст обработчика находится в репозитории в файле **Report_Список телефонов контактов.cls**


```vb
Private Sub Business_Phone_Click()
    If Not Me.Business_Phone.Value = "" Then
        Call ProstieZvonki.MakeCall(Me.Business_Phone.Value)
    End If
End Sub
```

Шаг 8. Добавим обработчик кнопки Позвонить на форме клиента через добавление макроса
------------------------------------------------------------------------------------

![Кнопка позвонить](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/call_button.png)

Создадим макрос для кнопки:

![Кнопка позвонить](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/button_macros.png)



Шаг 9. Умная переадресация
--------------------------


Чтобы проверить функцию трансфера, отправим запрос с помощью диагностической утилиты:

```
[events off]> Generate incoming 73430112233
```

В консоли сервера мы должны увидеть, что приложение отправило запрос на перевод звонка на нашего пользователя:

```
Transfer Event: callID = 18467, to = 101
```