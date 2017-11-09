Пример интеграции MS Access с сервисом Простые Звонки
==========================================================

Простые Звонки - сервис для интеграции клиентских приложений (Excel, 1C и ERP-cистем) с офисными и облачными АТС. Клиентское приложение может общаться с сервером Простых Звонков через единый API, независимо от типа используемой АТС. 

В данном примере мы рассмотрим процесс подключения к серверу Простых Звонков базы данных MS Access.

Мы возмём за основу стандартный шаблон MS Access под названием "Контакты" и добавим в него базовые функции:

- отображение всплывающей карточки при входящем звонке;
- звонок клиенту по клику на телефоный номер;
- умная переадресация на менеджера клиента;

Пример базы данных MS Access можно скачать из репозитория: [**ProstieZvonki-Contacts.accdb**](https://github.com/vedisoft/access-integration-tutorial/raw/master/ProstieZvonki-Contacts.accdb)


**Внимание!**

Не нужно копировать содержимое файлов .cls и .bas в свой проект. Эти файлы содержат дополнительные служебные данные и предназначены для импорта, а не для прямого копирования.
Так же, поскольку в путях и названиях используется кириллица, для корректной работы требуется, чтобы языком по умолчанию для не-юникодных(non-Unicode) символов был выбран Русский.

Шаг 0. Установка ActiveX
--------------------------

Необходимо скачать ActiveX по ссылке [отсюда](http://prostiezvonki.ru/installs/ProstieZvonki_ActiveX_2.0.exe)

После установки необходимо подключить ActiveX в редакторе VB кода в Tools -> References :

![Включаем ActiveX](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/references.png)

Шаг 1. Настройка подключения к тестовому серверу
--------------------------------------

Теперь нужно скачать [тестовый сервер и диагностическую утилиту](https://github.com/vedisoft/pz-developer-tools).

Запустим тестовый сервер:

    > TestServer.exe

и подключимся к нему диагностической утилитой:

    > Diagnostic.exe

    [events off]> Connect 127.0.0.1 asd
    Успешно начато установление соединения с АТС

Тестовое окружение настроено.

Шаг 2. Создадим VB класс для взаимодействия с ActiveX
------------------------------------------------------

Полный текст модуля, включающий обработку событий и загрузку настроек находится в репозитории в файле [**ProstieZvonkiWrapper.cls**](https://github.com/vedisoft/access-integration-tutorial/raw/master/ProstieZvonkiWrapper.cls)


Добавим метод для инициализации библиотеки и подключения:

```vb
Option Explicit

Public WithEvents prostie_zvonki_lib As CTIControlX

Const guidKey = "HKEY_CURRENT_USER\Software\Vedisoft\Access\GUID"
Const serverKey = "HKEY_CURRENT_USER\Software\Vedisoft\Access\ServerAddress"
Const passwordKey = "HKEY_CURRENT_USER\Software\Vedisoft\Access\ClientID"
Const phoneKey = "HKEY_CURRENT_USER\Software\Vedisoft\Access\Phone"

Public State As Boolean

Dim logPath As String
Dim ManagerPhone As String
Dim server As String
Dim password As String
Dim GUID As String

Public Sub Class_Initialize()
    On Error GoTo ErrorHandler
    CreateLogDir
    
    LoadSettings
    
    Set prostie_zvonki_lib = New CTIControlX
    Exit Sub
ErrorHandler:
    ShowError Err
    Resume Next
End Sub

Public Sub Connect()
    If (State) Then
        prostie_zvonki_lib.Disconnect
        State = False
    End If

    Dim ret As Integer
    On Error GoTo Errhandler
    prostie_zvonki_lib.phoneNumber = ManagerPhone
    ret = prostie_zvonki_lib.Connect(server, password, "Access", GUID, _
                                            logPath & "ProtocolLib_log.log", 0, 5000)
    If (ret = 0) Then
        State = True
    Else
        State = False
        MsgBox ("Can't connect to server")
    End If
    Exit Sub
Errhandler:
    ShowError Err
    Resume Next
End Sub
```

А также метод для совершения исходящего вызова, который будем использовать при нажатии на кнопку на форме:

```vb
Public Sub MakeCall(Phone As String)
    Call prostie_zvonki_lib.Call(ManagerPhone, Phone)
End Sub
```

В модуль также необходимо добавить вспомогательные методы для взаимодействия с реестром и генерации уникального идентификатора клиента ([см. полную версию модуля](https://github.com/vedisoft/access-integration-tutorial/raw/master/ProstieZvonkiWrapper.cls)).

Шаг 3. Добавим метод для обработки события входящего звонка
-----------------------------------------------------------

Для примера, будем показывать простое информационное сообщение, которое будет отображать имя звонящего (если клиент найден в базе данных) или номер телефона (если клиент не найден).
Имя будем искать по номеру телефона (колонка Рабочий телефон) в таблице Контакты.

```vb
Private Sub prostie_zvonki_lib_OnTransferredCall(ByVal CallID As String, _
                                                ByVal src As String, ByVal dst As String, ByVal line As String)
    On Error GoTo ErrorHandler
    
    If (dst <> ManagerPhone) Then
        Exit Sub
    End If
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT [Имя], [Фамилия] FROM Контакты" & _
            " WHERE Контакты.[Рабочий телефон] = '" & src & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        'show client name
        MsgBox ("Incoming call from " & rs.Fields(0).Value & " " & rs.Fields(1).Value)
    Else
        'can't find client, show number only
        MsgBox ("Incoming call from " & src)
    End If
    Call rs.Close
    Exit Sub
ErrorHandler:
    ShowError Err
    Resume Next
End Sub
```

![Входящий звонок](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/incoming_call.png)


Шаг 4. Добавим метод для обработки события OnTransferredRequest, используемый для умной переадресации
-----------------------------------------------------------------------------------------------------

```vb
Private Sub prostie_zvonki_lib_OnTransferRequest(ByVal CallID As String, ByVal from As String, ByVal line As String)
    On Error GoTo ErrorHandler
    Dim rs As Recordset
    Dim strSQL As String
    strSQL = "SELECT Менеджеры.[Телефон] FROM Менеджеры" & _
            " INNER JOIN Контакты ON Менеджеры.[Код] = Контакты.[Менеджер]" & _
            " WHERE Контакты.[Рабочий телефон] = '" & from & "'"
    Set rs = CurrentDb.OpenRecordset(strSQL)
    If rs.RecordCount >= 1 Then
        'call with number, if manager can handle call
        Call prostie_zvonki_lib.Transfer(CallID, rs.Fields(0).Value)
    Else
        'call with empty string, if manager can't handle call
        Call prostie_zvonki_lib.Transfer(CallID, "")
    End If
    Call rs.Close
    Exit Sub
ErrorHandler:
    ShowError Err
    Resume Next
End Sub
```

Для работы умной переадресации необходимо в типовой базе Контакты создать таблицу Менеджеры, содержащую ФИО и Внутренний номер телефона менеджера. В таблице Контакты создать колонку Менеджеры и связать эти таблицы между собой, чтобы иметь возможность назначать Ответственного менеджера для Клиента.

![Таблица Менеджеры](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/mangers_table.png)

А также создадим простую форму для создания Менеджеров

![Форма Менеджеры](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/managers.png)

Для того, чтобы иметь возможность установить для клиента ответственного менеджера, разместим на форме редактирования клиента выподающий список менеджеров

![ОТветственный менеджер](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/responsible_manager.png)

Шаг 5. Добавим метод для обработки события OnCompletedCall, используемого для создания истории звонков
------------------------------------------------------------------------------------------------------

```vb
Private Sub prostie_zvonki_lib_OnCompletedCall(ByVal CallID As String, ByVal src As String, ByVal dst As String, ByVal duration As Long, ByVal startTimestampString As String, ByVal endTimestampString As String, ByVal direction As Long, ByVal record As String, ByVal line As String)
    On Error GoTo ErrorHandler
    Dim manager As String, client As String
    If (direction = 1) Then
        client = dst
        manager = src
    Else
        client = src
        manager = dst
    End If
    startTimestampString = prostie_zvonki_lib.ConvertUtcToLocal(startTimestampString)
    endTimestampString = prostie_zvonki_lib.ConvertUtcToLocal(endTimestampString)
    Dim startTime, endTime
    startTime = DateAdd("s", CLng(startTimestampString), DateSerial(1970, 1, 1))
    endTime = DateAdd("s", CLng(endTimestampString), DateSerial(1970, 1, 1))
    Dim clientName As String
    clientName = FindClientByPhone(client)
    MsgBox ("Звонок завершен" & Chr(13) & "Телефон менеджера: " & manager & Chr(13) & "Клиент: " & clientName & Chr(13) & "Клиент(телефон): " & client & Chr(13) & "Начало звонка: " & startTime & Chr(13) & "Конец звонка: " & endTime & Chr(13) & "Продолжительность: " & duration / 1000 & " секунд")
    Exit Sub
ErrorHandler:
    ShowError Err
    Resume Next
End Sub
```

Здесь реализовано всплывающее окно по завершении разговора, но скорее всего ввы захотите сохранять историю в базу, чтобы иметь возможность посмотреть ее позднее.

Чтобы проверить работу истории звонков, отправим запрос с помощью диагностической утилиты:

	[events off]> generate history 123 45678 1394623140 1394623165 25377 out ""

Появится всплывающее окно с информацией о звонке.

Посмотреть синтаксис команды можно введя "hеlp generate history"

Шаг 6. Создадим VB модуль для доступа к объекту класса и его методам
---------------------------------------------------

Полный текст модуля находится в репозитории в файле [**ProstieZvonki.bas**](https://github.com/vedisoft/access-integration-tutorial/raw/master/ProstieZvonki.bas)

```vb
Option Compare Database
Option Explicit

Public prostie_zvonki_wrapper As ProstieZvonkiWrapper

Public Sub CreateWrapper()
    Set prostie_zvonki_wrapper = New ProstieZvonkiWrapper
End Sub

Public Function Init_Prostie_Zvonki(ManagerPhone As String)
    If (prostie_zvonki_wrapper Is Nothing) Then
        CreateWrapper
    End If
    prostie_zvonki_wrapper.SetManagerPhone (ManagerPhone)
End Function

Public Function MakeCall(phone As String)
    Call prostie_zvonki_wrapper.MakeCall(phone)
End Function
```

Шаг 7. Инициализация объекта взаимодействия с ActiveX
-----------------------------------------------------

Для инициализации нам необходимо определить какой именно менеджер работает с базой данных.
Для этого мы создадим простую форму Входа в систему, где будет выбираться текущий менеджер

![Форма входа](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/login.png)

На кнопку Подтвердить назначим макрос, который будет вызывать инициализацию

![Макрос инициализации](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/init_macros.png)

При успешном подключении мы должны увидеть в логе сервера:

	New client connected: 24DD18D4-C902-497F-A64B-28B2FA741661


Шаг 8. Добавим обработчик для гиперссылки номера телефона в списке клиентов через обработку события Click
-----------------------------------------------------------------------------------------------------------
![Ссылка](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/hiperlink.png)

Полный текст обработчика находится в репозитории в файле [**Report_Список телефонов контактов.cls**](https://github.com/vedisoft/access-integration-tutorial/raw/master/Report_%D0%A1%D0%BF%D0%B8%D1%81%D0%BE%D0%BA%20%D1%82%D0%B5%D0%BB%D0%B5%D1%84%D0%BE%D0%BD%D0%BE%D0%B2%20%D0%BA%D0%BE%D0%BD%D1%82%D0%B0%D0%BA%D1%82%D0%BE%D0%B2.cls)


```vb
Private Sub Business_Phone_Click()
    If Not Me.Business_Phone.Value = "" Then
        Call ProstieZvonki.MakeCall(Me.Business_Phone.Value)
    End If
End Sub
```

При клике по ссылке мы должны увидеть в логе сервера:
```
Call Event: from = 123, to = 73430112233
```

Шаг 9. Добавим обработчик кнопки Позвонить на форме клиента через добавление макроса
------------------------------------------------------------------------------------

![Кнопка позвонить](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/call_button.png)

Создадим макрос для кнопки:

![Кнопка позвонить](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/button_macros.png)

При нажатии на кнопку мы должны увидеть в логе сервера:

	Call Event: from = 123, to = 73430112233


Шаг 10. Всплывающая карточка
---------------------------

Для проверки всплывающей карточки отправим запрос с помощью диагностической утилиты:

	
	[events off]> Generate transfer 73430000000 123

Здесь 123 - внутренний номер текущего менеджера
Должна отобразиться всплывающая карточка, отображающая номер звонящего, т.к. номер не найден в базе.

Если мы выполним запрос, указав номер существующего клиента:

	 [events off]> Generate transfer 73430112233 123

Должна отобразиться всплывающая карточка с именем звонящего.

Шаг 11. Умная переадресация
--------------------------

Необходимо выбрать ответственного менеджера на карточке клиента.

Чтобы проверить функцию трансфера, отправим запрос с помощью диагностической утилиты:


	[events off]> Generate incoming 73430112233


В консоли сервера мы должны увидеть, что приложение отправило запрос на перевод звонка на нашего пользователя:

	Transfer Event: callID = 18467, to = 123

Значит система верно определила, что мы являемся ответственным сотрудником и хотим обслужить вызов.

Шаг 12. Настройки
-----------------
Добавим экран настроек, который будет обладать следующей функциональностью:

- поле для ввода ip адреса и пароля, который далее передаётся для инициализации
- статус подключения (свойство ActiveX ConnectionState)
- кнопки для подключения/отключения

Полный текст модуля обработчика окна настроек находится в репозитории в файле:
[**Form_Настройки.cls**](https://raw.github.com/vedisoft/access-integration-tutorial/master/Form_%D0%9D%D0%B0%D1%81%D1%82%D1%80%D0%BE%D0%B9%D0%BA%D0%B8.cls)

![Окно настроек](https://github.com/vedisoft/access-integration-tutorial/raw/master/img/acc_settings.png)
