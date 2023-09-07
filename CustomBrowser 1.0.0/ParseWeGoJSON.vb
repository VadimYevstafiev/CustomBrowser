Imports System.IO
Imports System.Text
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Json

Module ParseWeGoJSON
    Public Function ParseWeGoJSON(ByVal content As String) As String(,)
        Dim json As DataContractJsonSerializer = New DataContractJsonSerializer(GetType(HeadWeGoJSON))
        Dim data As HeadWeGoJSON = DirectCast(json.ReadObject(New MemoryStream(Encoding.UTF8.GetBytes(content))), HeadWeGoJSON)
        Dim result(data.WeGoResults.Count, 2) As String
        For i = 0 To data.WeGoResults.Count - 1
            result(i, 0) = data.WeGoResults(i).Title
            result(i, 1) = data.WeGoResults(i).Vicinity
        Next
        Return result
    End Function
End Module


<DataContract([Namespace]:="")>
Public Class HeadWeGoJSON
    <DataMember(Name:="results")>
    Public Property WeGoResults As List(Of WeGoResult)
End Class

<DataContract>
Public Class WeGoResult
    <DataMember(Name:="title")>
    Public Property Title As String

    <DataMember(Name:="highlightedTitle")>
    Public Property HighLightedTitle As String

    <DataMember(Name:="vicinity")>
    Public Property Vicinity As String

    <DataMember(Name:="highlightedVicinity")>
    Public Property HighLightedVicinity As String

    <DataMember(Name:="position")>
    Public Property Position As List(Of Double)

    <DataMember(Name:="category")>
    Public Property Category As String

    <DataMember(Name:="bbox")>
    Public Property Bbox As List(Of Double)

    <DataMember(Name:="href")>
    Public Property Href As String

    <DataMember(Name:="type")>
    Public Property Type As String

    <DataMember(Name:="id")>
    Public Property Id As String
End Class