Imports System.Threading
Imports System.Windows.Forms

Module MainProcess
    Private BrowserThread As Thread
    Private ArrayResult As String(,)
    Private StringResult As String
    Private IsStarted As Boolean
    Public streamWriter As New IO.StreamWriter("D:\Мои проекты\2.txt")

    Sub Main()
        Dim Message As String
        Dim AskMessage As String

        IsStarted = False
        Dim IsStoping As Boolean = False
        Do While IsStoping = False
            Console.WriteLine("Введите команду")
            Message = Console.ReadLine
            Select Case Message
                Case 1
                    Console.WriteLine("Введите идентификатор браузера")
                    BrowserID = Console.ReadLine
                    'Идентификатор браузера
                    '0 - Google
                    '1 - Yandex
                    '2 - WeGo
                    Start()
                    Do While IsStarted = False
                        Thread.Sleep(50)
                    Loop
                Case 2
                    Console.WriteLine("Введите выражение для запроса")
                    AskMessage = Console.ReadLine
                    Ask(AskMessage)
                    Thread.Sleep(2000)
                    Console.WriteLine(StringResult)

                Case 3
                    Console.WriteLine("Введите выражение для запроса и перехода")
                    AskMessage = Console.ReadLine
                    AskAndGo(AskMessage)
                    Thread.Sleep(2000)
                    Console.WriteLine(StringResult)
                Case 4
                    stopThread()
                    IsStoping = True
                Case Else
                    Exit Select
            End Select
        Loop
        Application.Exit()
    End Sub

    Delegate Function _CheckFullLoadPage()
    Private Sub Start()
        BrowserFormLoad = False
        BrowserStoped = False
        BrowserThread = New Thread(AddressOf startBrowserThread)
        BrowserThread.SetApartmentState(ApartmentState.STA)
        BrowserThread.Start()
        Do While BrowserFormLoad = False
            Thread.Sleep(50)
        Loop
        If Browser.InvokeRequired = False Then
            Do While Browser.InvokeRequired = False
                Thread.Sleep(50)
            Loop
        End If
        Dim metCheckFullLoadPage As New _CheckFullLoadPage(AddressOf CheckFullLoadPage)
        Do While Browser.Invoke(metCheckFullLoadPage, New Object() {}) = False
            Thread.Sleep(50)
        Loop
        Thread.Sleep(5000)
        IsStarted = True
    End Sub

    Delegate Sub _SendMessage(ByVal message As String)
    Delegate Function _CheckInput()
    Delegate Sub _ClearInput()
    Private Sub Ask(ByVal message As String)
        resultAsking = ""
        Dim result As String
        Dim input As String
        Dim messageLength = message.Length
        Dim metSendMessage As New _SendMessage(AddressOf SendMessage)
        Dim metClearInput As New _ClearInput(AddressOf ClearInput)
        Dim metCheckInput As New _CheckInput(AddressOf CheckInput)
        Dim litera(messageLength) As String
        For i As Integer = 1 To messageLength
            litera(i) = Mid(message, i, 1)
            input = Mid(message, 1, i)
            Browser.Invoke(metSendMessage, New Object() {litera(i)})
            result = Browser.Invoke(metCheckInput, New Object() {})
            Do Until input = result
                Thread.Sleep(50)
                result = Browser.Invoke(metCheckInput, New Object() {})
            Loop
        Next
        Do While resultAsking = ""
            Thread.Sleep(50)
        Loop
        Select Case BrowserID
            Case 1
                ArrayResult = ParseYandexJSON.ParseYandexJSON(resultAsking)
            Case 2
                ArrayResult = ParseWeGoJSON.ParseWeGoJSON(resultAsking)
        End Select
        Browser.Invoke(metClearInput, New Object() {})
        StringResult = ""
        For i = 0 To ArrayResult.GetUpperBound(0) - 1
            For j = 0 To ArrayResult.GetUpperBound(1) - 1
                StringResult = StringResult + ArrayResult(i, j) & "xАx"
            Next
        Next
    End Sub

    Delegate Function _CheckGoing()
    Private Sub AskAndGo(ByVal message As String)
        Dim result As String
        Dim input As String
        Dim messageLength = message.Length
        Dim metSendMessage As New _SendMessage(AddressOf SendMessage)
        Dim metClearInput As New _ClearInput(AddressOf ClearInput)
        Dim metCheckInput As New _CheckInput(AddressOf CheckInput)
        Dim metCheckGoing As New _CheckGoing(AddressOf CheckGoing)
        Dim litera(messageLength) As String
        For i As Integer = 1 To messageLength
            litera(i) = Mid(message, i, 1)
            input = Mid(message, 1, i)
            Browser.Invoke(metSendMessage, New Object() {litera(i)})
            result = Browser.Invoke(metCheckInput, New Object() {})
            Do Until input = result
                Thread.Sleep(50)
                result = Browser.Invoke(metCheckInput, New Object() {})
            Loop
        Next
        result = ""
        Browser.Invoke(metSendMessage, New Object() {"{ENTER}"})
        Do While result = ""
            result = Browser.Invoke(metCheckGoing, New Object() {})
            Thread.Sleep(50)
        Loop
        Dim x As Integer = result.IndexOf("map=")
        x = x + 5
        Dim z As Integer = x
        Dim y As Integer = 0
        Do While y < 2
            If Mid(result, z, 1) = "," Then
                y = y + 1
                z = z + 1
            Else
                z = z + 1
            End If
        Loop
        StringResult = Mid(result, x, z - x - 1)
        Thread.Sleep(2000)
        Browser.Invoke(metClearInput, New Object() {})
    End Sub

    Delegate Sub _StopBrowser()
    Private Sub stopThread()
        Dim metStopBrowser As New _StopBrowser(AddressOf StopBrowser)
        BrowserForm.Invoke(metStopBrowser, New Object() {})
        Do While BrowserStoped = False
            Thread.Sleep(50)
        Loop
        BrowserThread.Abort()
    End Sub
End Module
