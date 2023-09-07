Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json

Module ParseYandexJSON
    Public Function ParseYandexJSON(ByVal content As String) As String(,)
        Dim i As Integer = content.IndexOf("(") + 2
        Dim j As Integer = content.IndexOf(")")
        Dim subcontent As String = Mid(content, i, j - i + 1)
        streamWriter.WriteLine("-------------------------------------")
        streamWriter.WriteLine(subcontent)
        streamWriter.Flush()
        Dim json As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(HeadYandexJSON))
        Dim data As HeadYandexJSON = DirectCast(json.ReadObject(New MemoryStream(Encoding.UTF8.GetBytes(subcontent))), HeadYandexJSON)
        Dim result(data.YandexResults.Count, 2) As String
        For i = 0 To data.YandexResults.Count - 1
            result(i, 0) = data.YandexResults(i).Title.Text
            result(i, 1) = data.YandexResults(i).SubTitle.Text
        Next
        Return result
    End Function
End Module


<DataContract([Namespace]:="")>
Public Class HeadYandexJSON
    <DataMember(Name:="part")>
    Public Property Part As String

    <DataMember(Name:="suggest_reqid")>
    Public Property Suggest_Reqid As String

    <DataMember(Name:="results")>
    Public Property YandexResults As List(Of YandexResult)

End Class

<DataContract>
Public Class YandexResult
    <DataMember(Name:="type")>
    Public Property Type As String

    <DataMember(Name:="search_query")>
    Public Property Search_Query As String

    <DataMember(Name:="log")>
    Public Property Log As String

    <DataMember(Name:="title")>
    Public Property Title As TitleItems

    <DataMember(Name:="subtitle")>
    Public Property SubTitle As TitleItems

    <DataMember(Name:="text")>
    Public Property Text As String

    <DataMember(Name:="tags")>
    Public Property Tags As List(Of String)

    <DataMember(Name:="action")>
    Public Property Action As String

    <DataMember(Name:="need_loc")>
    Public Property Need_Loc As Integer?

    <DataMember(Name:="uri")>
    Public Property Uri As String

    <DataMember(Name:="distance")>
    Public Property Distance As DistanceItems
End Class

<DataContract>
Public Class LogItems
    <DataMember(Name:="type")>
    Public Property Type As String

    <DataMember(Name:="what")>
    Public Property What As WhatItem

    <DataMember(Name:="where")>
    Public Property Where As WhereItem
End Class

<DataContract>
Public Class WhatItem
    <DataMember(Name:="id")>
    Public Property Id As String
End Class

<DataContract>
Public Class WhereItem
    <DataMember(Name:="name")>
    Public Property Name As String
End Class

<DataContract>
Public Class TitleItems
    <DataMember(Name:="text")>
    Public Property Text As String

    <DataMember(Name:="hl")>
    Public Property Hl As List(Of List(Of Integer))
End Class

<DataContract>
Public Class DistanceItems
    <DataMember(Name:="value")>
    Public Property Value As Double

    <DataMember(Name:="text")>
    Public Property Text As String
End Class