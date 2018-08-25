Option Strict On
'**************************************************************************
' Name       :�H�i�}�X�^�R���N�V����
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/01/15
'        By  :FMS)Y.Maki
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MFd

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :�A�C�e���Q��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MFd
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MFd)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :�A�C�e���Q�Ɓi�H�i�}�X�^�j
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/05/16
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strFdCd As String) As Rec_MFd
        Try
            Key = Nothing
            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            For Each objMsg As Rec_MFd In InnerList
                If objMsg.FdCd.Trim = strFdCd.Trim Then
                    Key = objMsg
                    Exit For
                End If
            Next
            '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^

            Return Key

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '**********************************************************************
    ' Name       :�V�K���R�[�h���(�ėp)
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/02
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function NewRecord() As Object
        Dim objRecMFd As Rec_MFd

        objRecMFd = New Rec_MFd()
        Return objRecMFd
    End Function

    '**********************************************************************
    ' Name       :�A�C�e���ǉ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MFd) As Integer
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MFd)
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MFd

        Dim objCol As New Col_MFd
        '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
        For intIdx As Integer = 0 To Me.Count - 1
            ' Item���N���[�� �R�s�[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MFd))
        Next
        '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^
        Return objCol

    End Function

End Class