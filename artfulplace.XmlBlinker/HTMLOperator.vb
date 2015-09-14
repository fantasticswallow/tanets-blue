Imports System.Xml.Linq

Public Class HTMLOperator

    ''' <summary>
    ''' タイトル要素を作成します
    ''' </summary>
    ''' <param name="pageObj"></param>
    ''' <param name="addObj"></param>
    ''' <param name="lastModified"></param>
    ''' <returns></returns>
    Friend Shared Function AddTitleObject2(pageObj As XElement, addObj As XElement, lastModified As String) As XElement
        Dim titleValue = SpanEliminate(addObj.Value)
        Dim titleObj = <p class="SubTitle"></p>
        titleObj.Add(titleValue)
        titleObj.Add(<br/>)
        Dim lastModifiedObj = <span class="lastModified"></span>
        lastModifiedObj.Add(New XText(String.Format("last modified : {0}", lastModified)))
        titleObj.Add(lastModifiedObj)
        pageObj.Add(titleObj)
        pageObj.Add(<hr/>)
        'pageObj.<head>.<title>.First.Add(New XText(titleValue))
        HTMLGenerator.TitleArticle = titleValue.Value
        Return pageObj
    End Function


    ''' <summary>
    ''' CDATA 内の不要なspanを削除します
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Friend Shared Function SpanEliminate(value As String) As XElement
        If String.IsNullOrEmpty(value) Then
            Return <xemp/>
        End If
        Dim val2 = "<hoge>" + value + "</hoge>" 'Eliminate時の複数ルートオブジェクト指定回避
        Dim obj = XElement.Parse(val2)
        Dim retObj = <span></span>
        Dim lastStyle = ""
        For Each xd In obj.Nodes
            If xd.NodeType = Xml.XmlNodeType.Element Then
                Dim x = DirectCast(xd, XElement)
                If x.Name = "span" Then
                    If x.Attributes.Any(Function(xa) xa.Name = "style") Then
                        If x.Attributes.Any(Function(xa) xa.Name = "lang") Then
                            x.Attribute("lang").Remove()
                        End If

                        Dim styleStr = x.@style
                        Dim spl = styleStr.Split(";"c).Where(Function(xs) Not xs.StartsWith("font-family") AndAlso Not xs.StartsWith("font-size") AndAlso Not xs.Contains("background:white"))

                        Dim newStyle = ""
                        For Each spx In spl
                            newStyle += spx + ";"
                        Next

                        Dim nbsps = x.<nbsp>
                        For Each nb In nbsps.ToArray
                            nb.ReplaceWith(New XText(" "))
                        Next

                        If newStyle <> "" Then
                            If newStyle = lastStyle Then
                                DirectCast(retObj.LastNode, XElement).Add(x.Value)
                            Else
                                x.@style = newStyle
                                retObj.Add(x)
                            End If
                        Else
                            lastStyle = ""
                            retObj.Add(x.Value)
                        End If
                    Else
                        lastStyle = ""
                        retObj.Add(x.Value)
                    End If
                Else
                    lastStyle = ""
                    retObj.Add(x)
                End If
            ElseIf xd.NodeType = Xml.XmlNodeType.Text
                lastStyle = ""
                retObj.Add(xd)
            Else
                lastStyle = ""
            End If
        Next

        Return retObj
    End Function

    ''' <summary>
    ''' 元のstringからXElement構築に不要なデータを削除します
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Friend Shared Function BaseReplace(str As String)
        str = str.Replace("<one:", "<").Replace("</one:", "</") 'OneNoteのネームスペース回避
        str = str.Replace("&nbsp;", "<nbsp/>") 'nbspの消滅対策
        'mathMLのネームスペースその他回避
        str = str.Replace("<mml:", "<").Replace("</mml:", "</").Replace("xmlns:mml=""http://www.w3.org/1998/Math/MathML""", "xmlns=""http://www.w3.org/1998/Math/MathML""").Replace("[<!--[if mathML]>", "[").Replace("<![endif]-->]", "]")
        Return str.Replace("lang=en-US", "lang=""en-US""").Replace("lang=ja", "lang=""ja""").Replace("<br>", "<br/>") 'XElementの仕様
    End Function
End Class
