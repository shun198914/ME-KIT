Option Strict On
'**************************************************************************
' Name       :�X�g�A�h�}�X�^�R���N�V����
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :
'        By  :
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MStrd
    Inherits CollectionBase
    '**********************************************************************
    ' Name       :�A�C�e���Q��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MStrd
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MStrd)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :�A�C�e���Q�Ɓi�p�����[�^���j
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/06/30
    '        By  :Y.Miyoshi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strParameterName As String) As Rec_MStrd
        Try
            Key = Nothing
            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            For Each objMsg As Rec_MStrd In InnerList
                If objMsg.ParameterName.Trim = strParameterName.Trim Then
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
    ' Name       :�A�C�e���ǉ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MStrd) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :�A�C�e���ǉ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal ParameterName As String, ByVal SqlDbType As Data.SqlDbType, ByVal Size As Integer, ByVal Value As String) As Integer
        Try
            InnerList.Add(New Rec_MStrd(ParameterName, SqlDbType, Size, Value))
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
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MStrd)
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
    ' Design On  :
    '        By  :
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
End Class
