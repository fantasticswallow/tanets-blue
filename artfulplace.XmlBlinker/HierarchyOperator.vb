Public Class HierarchyOperator
    Public Shared Function FilterSections(xs As String) As IEnumerable(Of Tuple(Of String, String, String, String))
        xs = xs.Replace("<one:", "<").Replace("</one:", "</") 'OneNoteのネームスペース回避
        Dim xe = XElement.Parse(xs)
        Dim xnt = xe.<Notebook>

        Return xnt.SelectMany(Function(x) FilterFromNotebook(x, x.@name, x.@ID, New List(Of Tuple(Of String, String, String, String))))
    End Function

    Private Shared Function FilterFromNotebook(xe As XElement, ntName As String, nId As String, ls As List(Of Tuple(Of String, String, String, String))) As List(Of Tuple(Of String, String, String, String))
        For Each xl In xe.Elements
            If xl.Name = "Section" Then
                ls.Add(Tuple.Create(ntName, nId, xl.@name, xl.@ID))
            ElseIf xl.Name = "SectionGroup" Then
                ls = FilterFromNotebook(xl, ntName, nId, ls)
            End If
        Next
        Return ls
    End Function

End Class
