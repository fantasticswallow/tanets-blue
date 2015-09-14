
Imports System.Globalization
Imports System.Xml
''' <summary>
''' XElementを用いたHTML Generaterを提供します
''' </summary>
''' <remarks></remarks>
Public Class HTMLGenerator

    ''' <summary>
    ''' OneNote XMLをもとに近いHTML要素を生成します
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Shared Function CreateFromXml(str As String) As XElement
        str = HTMLOperator.BaseReplace(str)

        Dim xe = XElement.Parse(str)

        HTMLGenerator.PageId = xe.@ID
        Dim lastMod = DateTime.ParseExact(xe.@lastModifiedTime, "yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.CurrentCulture)

        'lastModifiedTime="2015-09-04T10:17:48.000Z"

        Dim pageObj = <div></div>
        For Each xn In xe.Elements()
            If xn.Name = "Title" Then
                pageObj = HTMLOperator.AddTitleObject2(pageObj, xn, lastMod.ToString("yyyy/MM/dd HH:mm:ss"))
            ElseIf xn.Name = "Outline" Then
                pageObj.Add(HTMLGenerator.CreateOutlineObject(xn))
            ElseIf xn.Name = "Image"
                pageObj.Add(HTMLGenerator.CreateImageObject(xn, PageId))
            End If
        Next

        Return pageObj
    End Function

    ''' <summary>
    ''' Outline要素を変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Friend Shared Function CreateOutlineObject(obj As XElement) As XElement
        ' Positionを保持しない？

        Dim pos = obj.<Position>
        Dim x = pos.@x
        Dim y = pos.@y
        Dim ydb = Double.Parse(y)
        ydb -= 90
        Dim stylestr = String.Format("margin-left: {0}px; margin-top: {1}px;", x, ydb)
        Dim rootEle = <div style=<%= stylestr %> class=<%= divStyle() %>></div>

        'Dim rootEle = <div></div>

        For Each xe In obj.<OEChildren>.Elements
            Dim cres = CreateOEObject(xe)
            If cres.NodeType <> Xml.XmlNodeType.Element OrElse Not DirectCast(cres, XElement).Name = "xemp" Then
                rootEle.Add(cres)
            End If
            rootEle.Add(<br/>)
            rootEle.Add(New XText(vbCrLf))
        Next

        Return rootEle
    End Function

    ''' <summary>
    ''' T要素を変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Friend Shared Function CreateOETextObject(obj As XElement) As XElement
        Dim text = obj.<T>
        Dim res2 = <span></span>
        For Each tx In text
            Dim res = HTMLOperator.SpanEliminate(tx.Value)
            If Not res.Name = "xemp" Then
                res2.Add(res)
            End If
        Next

        If res2.Elements.Count = 1 Then
            Return res2.<span>.First
        ElseIf res2.Elements.Count = 0
            Return <xemp/>
        Else
            Return res2
        End If
    End Function

    ''' <summary>
    ''' Table要素を変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Friend Shared Function CreateOETableObject(obj As XElement) As XElement
        Dim root2 = <div class="doctable"></div>
        Dim rootObj = <table></table>
        Dim tb = obj.<Table>.<Row>
        For Each r In tb
            rootObj.Add(CreateTableRowObject(r))
        Next

        root2.Add(rootObj)
        Return root2
    End Function

    ''' <summary>
    ''' TableのRow要素を変換します。
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Friend Shared Function CreateTableRowObject(obj As XElement) As XElement
        Dim rootObj = <tr></tr>
        Dim cells = obj.<Cell>
        For Each c In cells
            Dim td = <td></td>
            Dim res = HTMLOperator.SpanEliminate(c.Value)
            If Not res.Name = "xemp" Then
                td.Add(res)
            End If
            rootObj.Add(td)
        Next
        Return rootObj
    End Function

    ''' <summary>
    ''' OE要素を変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <returns></returns>
    Friend Shared Function CreateOEObject(obj As XElement) As XNode
        Dim tn = obj.Elements.First.Name
        Select Case tn
            Case "T"
                Dim textEle = CreateOETextObject(obj)
                If obj.Attributes.Any(Function(xa) xa.Name = "style") Then
                    Dim styleStr = obj.@style
                    Dim spl = styleStr.Split(";"c).Where(Function(xs) Not xs.StartsWith("font-family") AndAlso Not xs.StartsWith("font-size"))
                    Dim newStyle = ""
                    For Each spx In spl
                        newStyle += spx + ";"
                    Next
                    If newStyle <> "" Then
                        textEle.@style = newStyle
                    End If
                    'If styleStr.Contains("Consolas") Then 'Consolasをもとにソースコードを識別する(Visual Studioからのコピー専用)
                    '    Dim preEle = <pre></pre>
                    '    preEle.Add(textEle)
                    '    Return preEle
                    'End If
                End If
                If textEle.Attributes.Count = 0 AndAlso textEle.Nodes.Count = 1 AndAlso textEle.Nodes.First.NodeType = Xml.XmlNodeType.Text Then
                    Return New XText(textEle.Value)
                End If
                Return textEle
            Case "Table"
                Return CreateOETableObject(obj)
            Case "Image"
                Return CreateImageObjectInOE(obj, PageId)
            Case Else
                Return <span></span>
        End Select
    End Function

    ''' <summary>
    ''' Pageオブジェクト内にあるImageオブジェクトを変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="pId"></param>
    ''' <returns></returns>
    Friend Shared Function CreateImageObject(obj As XElement, pId As String) As XElement
        Dim pos = obj.<Position>
        Dim x = pos.@x
        Dim y = pos.@y
        Dim stylestr = String.Format("margin-left: {0}px; margin-top: {1}px;", x, y)
        Dim rootEle = <div style=<%= stylestr %> class=<%= divStyle() %>></div>

        Dim dataFormat = String.Format("data:image/{0};base64,", obj.@format)
        dataFormat += GetBinaryObject(pId, obj.<CallbackID>.@callbackID)


        Dim imgEle = <img src=<%= dataFormat %> height=<%= obj.<Size>.@height %> width=<%= obj.<Size>.@width %>/>

        rootEle.Add(imgEle)

        Return rootEle
    End Function

    ''' <summary>
    ''' OE内のImageオブジェクトを変換します
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="pId"></param>
    ''' <returns></returns>
    Friend Shared Function CreateImageObjectInOE(obj As XElement, pId As String) As XElement
        Dim dataFormat = String.Format("data:image/{0};base64,", obj.<Image>.@format)
        dataFormat += GetBinaryObject(pId, obj.<Image>.<CallbackID>.@callbackID)


        Dim imgEle = <img src=<%= dataFormat %> height=<%= obj.<Image>.<Size>.@height %> width=<%= obj.<Image>.<Size>.@width %>/>


        Return imgEle
    End Function

    ''' <summary>
    ''' Imageのデータを取得するための関数
    ''' </summary>
    Public Shared GetBinaryObject As Func(Of String, String, String)
    ''' <summary>
    ''' 現在のページのID
    ''' </summary>
    Public Shared PageId As String

    ''' <summary>
    ''' ページのタイトル
    ''' </summary>
    Public Shared TitleArticle As String

    ''' <summary>
    ''' フォントサイズを変更するか
    ''' </summary>
    Public Shared IsResizeFont As Boolean

    ''' <summary>
    ''' XML Declarationを書き込まずにXElementを保存します。
    ''' </summary>
    ''' <param name="path"></param>
    ''' <param name="xe"></param>
    Public Shared Sub WriteXml(path As String, xe As XElement)
        Dim xws As XmlWriterSettings = New XmlWriterSettings()
        xws.OmitXmlDeclaration = True
        xws.Indent = True
        xws.IndentChars = "  "

        Using writer = XmlWriter.Create(path, xws)
            xe.Save(writer)
        End Using
    End Sub

    ''' <summary>
    ''' IsResizeFontに対応するstyleを返します
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function divStyle() As String
        If IsResizeFont Then
            Return "style3"
        Else

            Return "style1"
        End If
    End Function

End Class
