Imports System.IO
Imports System.Text

Public Class JSONDecodeException
    Inherits Exception
    Sub New(ByVal strMsg As String)
        MyBase.New(strMsg)
    End Sub
End Class

Public Class JSONEncodeException
    Inherits Exception
    Sub New(ByVal strMsg As String)
        MyBase.New(strMsg)
    End Sub
End Class

Public NotInheritable Class JSON
    Private Sub New()

    End Sub

    ''' <summary>
    ''' 指定のファイルを読み込み、全文を取得します
    ''' </summary>
    ''' <param name="fp"></param>
    ''' <returns></returns>
    Private Shared Function ReadJson(ByVal fp As String) As String
        Using r As New StreamReader(fp)
            Return r.ReadToEnd()
        End Using
    End Function

    ''' <summary>
    ''' 文字列からJson読み込み
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns></returns>
    Public Shared Function Loads(data As String) As Object
        Return 1
    End Function

    Private Shared Function GetValueTypeFromString(ByVal strValue As String) As Type
        Dim typeValue As Type = Nothing
        If strValue.StartsWith(""""c) Then
            typeValue = GetType(String)
        ElseIf strValue.StartsWith("["c) Then
            typeValue = GetType(List(Of))
        ElseIf strValue.StartsWith("{"c) Then
            typeValue = GetType(Dictionary(Of ,))
        ElseIf strValue = "true" OrElse strValue = "false" Then
            typeValue = GetType(Boolean)
        ElseIf strValue = "null" Then
            typeValue = Nothing
        ElseIf Double.TryParse(strValue, 0) Then
            typeValue = If(strValue.Contains("."c), GetType(Double), GetType(Integer))
        Else
            Throw New JSONDecodeException("Type Error")
        End If
        Return typeValue
    End Function
    Private Shared Function GetValueFromString(ByVal strValue As String) As Object
        Dim value As Object
        Select Case GetValueTypeFromString(strValue)
            Case GetType(String)
                value = strValue.Trim(""""c).Replace("\"c, "")
            Case GetType(Boolean)
                value = If(strValue = "true", True, False)
            Case GetType(Double)
                value = Double.Parse(strValue)
            Case GetType(Integer)
                value = Integer.Parse(strValue)
            Case GetType(List(Of))
                value = ParseList(strValue & "]"c)
            Case GetType(Dictionary(Of,))
                value = ParseDictionary(strValue & "}"c)
            Case Nothing
                value = Nothing
        End Select
        Return value
    End Function

    Private Shared Function ParseDictionary(ByVal strJson As String) As Dictionary(Of String, Object)
        If Not (strJson.StartsWith("{"c)) Then Throw New Exception()

        Dim dctRoot As New Dictionary(Of String, Object)
        Dim blnInList As Boolean = False
        Dim blnInKey As Boolean = False
        Dim strKeyBuilder As New StringBuilder
        Dim key As String = String.Empty
        Dim strValueBuilder As New StringBuilder
        Dim privChar As Char

        For Each c As Char In strJson
            If Not blnInKey Then
                Select Case c
                    Case "["c
                        blnInList = True
                    Case "]"c
                        blnInList = False
                End Select
            End If

            Select Case c
                Case "{"c
                Case """"c
                    If privChar <> "\"c Then blnInKey = If(blnInKey, False, True)
                Case ":"c
                    If Not blnInKey AndAlso key = String.Empty Then
                        key = strKeyBuilder.ToString().Trim().Trim(""""c).Replace("\"c, "")
                        strKeyBuilder.Clear()
                    End If
                Case ","c, "}"c, "]"c
                    If key <> String.Empty AndAlso Not blnInList Then
                        Dim strValue As String = strValueBuilder.ToString().Trim().Trim(":"c).Trim()
                        Dim value As Object = GetValueFromString(strValue)

                        dctRoot.Add(key, value)
                        key = String.Empty
                        strValueBuilder.Clear()
                    End If
                Case <> " "c
                    privChar = c
            End Select

            If blnInKey AndAlso key = String.Empty Then strKeyBuilder.Append(c)
            If key <> String.Empty Then strValueBuilder.Append(c)
        Next

        Return dctRoot
    End Function

    Private Shared Function ParseList(ByVal strJson As String) As List(Of Object)
        If Not (strJson.StartsWith("["c)) Then Throw New Exception()
        Dim lstRoot As New List(Of Object)

        Dim blnInDict As Boolean = False
        Dim blnInString As Boolean = False
        Dim strValueBuilder As New StringBuilder
        Dim privChar As Char

        For Each c As Char In strJson
            If Not blnInString Then
                Select Case c
                    Case "{"c
                        blnInDict = True
                    Case "}"c
                        blnInDict = False
                End Select
            End If

            Select Case c
                Case "["c
                Case "{"c
                    strValueBuilder.Append(c)
                Case """"c
                    If privChar <> "\"c Then blnInString = If(blnInString, False, True)
                    strValueBuilder.Append(c)
                Case ","c, "}"c, "]"c
                    If c = "}" OrElse (blnInDict AndAlso c = ",") Then strValueBuilder.Append(c)
                    If Not blnInString AndAlso Not blnInDict AndAlso strValueBuilder.Length <> 0 Then
                        Dim strValue As String = strValueBuilder.ToString().Trim()
                        Dim value As Object = GetValueFromString(strValue)
                        lstRoot.Add(value)
                        strValueBuilder.Clear()
                    End If
                Case Else
                    strValueBuilder.Append(c)
                    If c <> " "c Then privChar = c
            End Select
        Next
        Return lstRoot
    End Function

    Public Shared Sub ShowJson(ByVal dctData As Dictionary(Of String, Object), Optional ByVal depth As Integer = 0)
        Dim indent As String = New String(" "c, 4 * depth)
        Console.WriteLine("{")
        For i As Integer = 0 To dctData.Keys.Count - 1
            Dim strKey As String = dctData.Keys(i)
            Dim objValue As Object = dctData(strKey)
            Console.Write($"{New String(" "c, 4 * (depth + 1))}""{strKey}"": ")
            If TypeOf objValue Is Dictionary(Of String, Object) Then
                ShowJson(DirectCast(objValue, Dictionary(Of String, Object)), depth + 1)
            ElseIf TypeOf objValue Is List(Of Object) Then
                ShowList(DirectCast(objValue, List(Of Object)), depth + 1)
            Else
                Dim strValue As String = If(objValue IsNot Nothing, If(TypeOf objValue Is String, $"""{objValue}""", objValue.ToString()), "Nothing")
                Console.Write(strValue)
            End If
            Console.WriteLine($"{If(i <> dctData.Keys.Count - 1, ",", "")}")
        Next
        Console.Write($"{indent}{"}"}{If(depth = 0, vbCrLf, "")}")
    End Sub
    Public Shared Sub ShowList(ByVal lstData As List(Of Object), Optional ByVal depth As Integer = 0)
        Dim indent As String = New String(" "c, 4 * depth)
        Console.WriteLine("[")
        For i As Integer = 0 To lstData.Count - 1
            Dim objValue As Object = lstData(i)
            Console.Write($"{New String(" "c, 4 * (depth + 1))}")
            If TypeOf objValue Is Dictionary(Of String, Object) Then
                ShowJson(DirectCast(objValue, Dictionary(Of String, Object)), depth + 1)
            ElseIf TypeOf objValue Is List(Of Object) Then
                ShowList(DirectCast(objValue, List(Of Object)), depth + 1)
            Else
                Dim strValue As String = If(objValue IsNot Nothing, If(TypeOf objValue Is String, $"""{objValue}""", objValue.ToString()), "Nothing")
                Console.Write(strValue)
            End If
            Console.WriteLine(If(i <> lstData.Count - 1, ",", ""))
        Next
        Console.Write($"{indent}{"]"}{If(depth = 0, vbCrLf, "")}")
    End Sub

    ''' <summary>
    ''' ファイルからJson読み込み
    ''' </summary>
    ''' <param name="fp"></param>
    ''' <returns></returns>
    Public Shared Async Function Load(path As String) As Task(Of Object)
        Dim strJson As String = ReadJson(path)
        Dim dct As Dictionary(Of String, Object) = ParseDictionary(strJson)

        Console.WriteLine($"原本：{vbCrLf}{strJson}")
        ShowJson(dct)

        Dim result As Integer = Await Task.Run(
        Function() As Integer
            For index = 0 To Integer.MaxValue - 1
            Next
            Return 1
        End Function
     )
        Return result
    End Function

    ''' <summary>
    ''' データからJson文字列を作成
    ''' </summary>
    ''' <param name="data"></param>
    ''' <returns></returns>
    Public Shared Function Dumps(data As Dictionary(Of Object, Object)) As String
        Return ""
    End Function

    ''' <summary>
    ''' データからJsonファイルを出力
    ''' </summary>
    ''' <param name="path"></param>
    Public Shared Async Sub Dump(path As String)

    End Sub
End Class