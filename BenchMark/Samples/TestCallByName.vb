Imports System
Imports System.Reflection

Public Class TestCallByNameClass
#Region "�y�����o�ϐ��z"
    Private m_A As String = "A"
    Private m_B As String = "B"
    Private m_C As String = "C"
    Private m_D As String = "D"
    Private m_E As String = "E"
#End Region
#Region "�y�v���p�e�B�z"
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
#Region "�y���\�b�h�z"
    ''' <summary>
    ''' CallByName���g�p���ăv���p�e�B���Ăяo��
    ''' </summary>
    Public Sub CallMembersWithCallByName()
        For Each m In New String() {"A", "B", "C", "D", "E"}
            CallByName(Me, m, CallType.Get)
        Next
    End Sub

    ''' <summary>
    ''' �ʏ�ʂ�Ƀv���p�e�B���Ăяo��
    ''' </summary>
    Public Function CallMembers(i As Integer, j As Integer) As Boolean
        For Each m In New String() {A, B, C, D, E}  ' �������ŌĂяo��

        Next
        Return True
    End Function
#End Region
End Class

Module TestCallByName
    Sub Test()
        ' ������
        Dim times As Long = 1000000

        ' �e�X�g�N���X
        Dim clsTest As New TestCallByNameClass()

        ' �x���`�}�[�N(��������(s)���󂯎��)
        Dim CallByNameResult As clsBenchMark = clsBenchMark.Method(times, NameOf(clsTest.CallMembersWithCallByName), Sub() clsTest.CallMembersWithCallByName())
        Dim BetaGakiResult As clsBenchMark = clsBenchMark.Method(times, NameOf(clsTest.CallMembers), Sub() clsTest.CallMembers(0, 1))

        If (CallByNameResult < BetaGakiResult) Then
            Console.WriteLine($"{CallByNameResult.MethodName}������")
        ElseIf (CallByNameResult > BetaGakiResult) Then
            Console.WriteLine($"{BetaGakiResult.MethodName}������")
        Else
            Console.WriteLine("�����x")
        End If

    End Sub


End Module
