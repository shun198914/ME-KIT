Option Strict On
'**************************************************************************
' Name       :���M�E�E�ۉ��H�L�^�뗿���R���N�V����
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_RHVirusF

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :�A�C�e���Q��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_RHVirusF
        Get
            Try
                Return CType(InnerList.Item(index), Rec_RHVirusF)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :�A�C�e���ǉ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_RHVirusF) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :�A�C�e���}��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_RHVirusF)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Name       :�A�C�e���폜
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Remove(ByVal Index As Integer)
        Try
            InnerList.RemoveAt(Index)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Method     :�A�C�e���N���[��
    '
    '----------------------------------------------------------------------
    ' Comment    :����J�����o�Ƃ��� 
    '
    '----------------------------------------------------------------------
' Design On  :2008/04/22
'        By  :M.Okazawa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_RHVirusF

        Dim objCol As New Col_RHVirusF
        '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
        For intIdx As Integer = 0 To Me.Count - 1
            ' Item���N���[�� �R�s�[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_RHVirusF))
        Next
        '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^
        Return objCol

    End Function

End Class