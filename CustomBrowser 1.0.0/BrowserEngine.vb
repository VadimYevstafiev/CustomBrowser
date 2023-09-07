Imports Gecko
Imports Gecko.Cache
Imports Gecko.Events
Imports System.Drawing
Imports System.Windows.Forms

Module BrowserEngine
    Public WithEvents Browser As GeckoWebBrowser
    Public BrowserID As Integer
    Public DOM_Ready As Boolean = False
    Dim Url() As String = {"https://www.google.com.ua/maps/", "https://yandex.ru/maps/", "https://wego.here.com"}
    Public BrowserForm As Form
    Public BrowserFormLoad As Boolean
    Private WithEvents myObserver As New XMLRequestChecker
    Public resultAsking As String = ""
    Public BrowserStoped As Boolean

    Public Sub InitializeFirefox()
        Xpcom.EnableProfileMonitoring = False
        Xpcom.Initialize("Firefox")

        Dim CookieMan As nsICookieManager
        CookieMan = Xpcom.GetService(Of nsICookieManager)("@mozilla.org/cookiemanager;1")
        CookieMan = Xpcom.QueryInterface(Of nsICookieManager)(CookieMan)
        CookieMan.RemoveAll()
        ImageCache.ClearCache(True)
        ImageCache.ClearCache(False)
        GeckoPreferences.User("browser.cache.disk.enable") = False
        GeckoPreferences.User("places.history.enabled") = False

        GeckoPreferences.User("general.useragent.override") = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.59 Safari/537.36 OPR/41.0.2353.46"
        GeckoPreferences.User("general.appname.override") = "Netscape"
        GeckoPreferences.User("general.appversion.override") = "5.0 (Windows)"
        GeckoPreferences.User("general.oscpu.override") = "Windows NT 6.1"
        GeckoPreferences.User("general.platform.override") = "Win32"
        GeckoPreferences.User("general.productSub.override") = ""
        GeckoPreferences.User("general.buildID.override") = 0
        GeckoPreferences.User("general.useragent.vendor") = ""
        GeckoPreferences.User("general.useragent.vendorSub") = ""
        GeckoPreferences.User("intl.accept_languages") = "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"

        If BrowserID = 1 Then
            GeckoPreferences.User("network.proxy.type") = 1
            GeckoPreferences.User("network.proxy.http") = "37.29.76.238"
            GeckoPreferences.User("network.proxy.http_port") = 3128
            GeckoPreferences.User("network.proxy.ssl") = "37.29.76.238"
            GeckoPreferences.User("network.proxy.ssl_port") = 3128
        End If
    End Sub

    Public Sub startBrowserThread()
        BrowserFormLoad = False
        BrowserForm = New Form
        AddHandler BrowserForm.Load, AddressOf handlesBrowserFormLoad
        BrowserForm.Size = New Size(720, 560)
        BrowserForm.Location = New Point(50, 50)

        InitializeFirefox()
        Browser = New GeckoWebBrowser
        Browser.Dock = DockStyle.Fill
        BrowserForm.Controls.Add(Browser)

        DOM_Ready = False
        AddHandler Browser.DocumentCompleted, AddressOf Browser_DocumentCompleted
        If BrowserID = 1 Or BrowserID = 2 Then
            ObserverService.AddObserver(myObserver)
        End If

        Browser.Navigate(Url(BrowserID))
        Application.Run(BrowserForm)
    End Sub

    Private Sub handlesBrowserFormLoad(sender As Object, e As EventArgs)
        BrowserFormLoad = True
    End Sub

    Private Sub Browser_DocumentCompleted(ByVal sender As Object, e As GeckoDocumentCompletedEventArgs)
        If e.Uri.AbsolutePath = TryCast(sender, GeckoWebBrowser).Url.AbsolutePath And TryCast(sender, GeckoWebBrowser).Document.ReadyState = "complete" Then
            DOM_Ready = True
        Else
            Return
        End If
    End Sub

    Public Function CheckFullLoadPage() As Boolean
        Dim result As String = ""
        Dim context = New AutoJSContext(Browser.Window)
        Select Case BrowserID
            Case 0
                context.EvaluateScript("document.querySelector('div#itamenu');", result)
            Case 1
                context.EvaluateScript("document.querySelector('div.home-panel-content-view__body');", result)
            Case 2
                context.EvaluateScript("document.querySelector('div.type_content');", result)
        End Select

        If DOM_Ready = False Then
            Return False
        Else
            If result = "null" Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Public Sub SendMessage(ByVal message As String)
        BrowserForm.Activate()
        Browser.Focus()
        Dim context = New AutoJSContext(Browser.Window)
        If BrowserID = 2 Then
            context.EvaluateScript("document.querySelector('#searchbar input').focus();")
        End If
        SendKeys.Send(message)
    End Sub

    Public Function CheckGoing() As String
        Dim result As String = ""
        Dim context = New AutoJSContext(Browser.Window)
        context.EvaluateScript("document.querySelector('div.address');", result)
        If result = "null" Then
            Return ""
        Else
            Return Browser.Url.AbsoluteUri
        End If
    End Function

    Public Function CheckInput() As String
        Dim result As String = ""
        Dim context = New AutoJSContext(Browser.Window)
        Select Case BrowserID
            Case 0
                context.EvaluateScript("document.getElementById('searchboxinput').value;", result)
            Case 1
                context.EvaluateScript("document.querySelector('.input_air-search-large__control').value;", result)
            Case 2
                context.EvaluateScript("document.querySelector('#searchbar input').value;", result)
        End Select

        Return result
    End Function

    Public Sub ReadResponseStream(ByVal content As String) Handles myObserver.CaptureResponseData
        resultAsking = content
        MessageBox.Show(resultAsking)
    End Sub

    Public Sub ClearInput()
        Browser.Focus()
        Dim context = New AutoJSContext(Browser.Window)
        Select Case BrowserID
            Case 0
                context.EvaluateScript("document.getElementById('searchboxinput').value='';")
            Case 1
                context.EvaluateScript("document.querySelector('.input_air-search-large__control').value='';")
            Case 2
                context.EvaluateScript("document.querySelector('#searchbar input').value='';")
        End Select

    End Sub

    Public Sub StopBrowser()
        Browser.Navigate("about:blank")
        Browser.Document.Cookie = ""
        Browser.Stop()
        BrowserForm.Close()
        BrowserStoped = True
    End Sub
End Module
