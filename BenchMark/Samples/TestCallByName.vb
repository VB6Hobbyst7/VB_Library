Imports System
Imports System.Reflection

Public Class TestCallByNameClass
#Region "【メンバ変数】"
    Private m_A As String = "A"
    Private m_B As String = "B"
    Private m_C As String = "C"
    Private m_D As String = "D"
    Private m_E As String = "E"
#End Region
#Region "【プロパティ】"
    Public ReadOnly Property A() As String
        Get
            Return m_A
        End Get
    End Property
    Public ReadOnly Property B() As String
        Get
            Return m_B
        End Get
    End Property
    Public ReadOnly Property C() As String
        Get
            Return m_C
        End Get
    End Property
    Public ReadOnly Property D() As String
        Get
            Return m_D
        End Get
    End Property
    Public ReadOnly Property E() As String
        Get
            Return m_E
        End Get
    End Property
#End Region
#Region "【メソッド】"
    ''' <summary>
    ''' CallByNameを使用してプロパティを呼び出す
    ''' </summary>
    Public Sub CallMembersWithCallByName()
        For Each m In New String() {"A", "B", "C", "D", "E"}
            CallByName(Me, m, CallType.Get)
        Next
    End Sub

    ''' <summary>
    ''' 通常通りにプロパティを呼び出す
    ''' </summary>
    Public Function CallMembers(i As Integer, j As Integer) As Boolean
        For Each m In New String() {A, B, C, D, E}  ' ←ここで呼び出し

        Next
        Return True
    End Function
#End Region
End Class

Module TestCallByName
    Sub Test()
        ' 処理回数
        Dim times As Long = 1000000

        ' テストクラス
        Dim clsTest As New TestCallByNameClass()

        ' ベンチマーク(処理時間(s)を受け取る)
        Dim CallByNameResult As clsBenchMark = clsBenchMark.Method(times, NameOf(clsTest.CallMembersWithCallByName), Sub() clsTest.CallMembersWithCallByName())
        Dim BetaGakiResult As clsBenchMark = clsBenchMark.Method(times, NameOf(clsTest.CallMembers), Sub() clsTest.CallMembers(0, 1))

        If (CallByNameResult < BetaGakiResult) Then
            Console.WriteLine($"{CallByNameResult.MethodName}が速い")
        ElseIf (CallByNameResult > BetaGakiResult) Then
            Console.WriteLine($"{BetaGakiResult.MethodName}が速い")
        Else
            Console.WriteLine("同速度")
        End If

    End Sub


End Module
