Option Strict On
'**************************************************************************
' Name       :�����}�X�^�R���N�V����
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2007/12/03
'        By  :CM)Kuwahara
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MNutri

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :�A�C�e���Q��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/03
    '        By  :CM)Kuwahara
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MNutri
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MNutri)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :�A�C�e���Q�Ɓi�����敪�j
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/15
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal intNtDiv As Integer) As Rec_MNutri
        Try
            Key = Nothing
            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            For Each objMsg As Rec_MNutri In InnerList
                If objMsg.NtDiv = intNtDiv Then
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
    ' Name       :�A�C�e���Q�Ɓi�����R�[�h�j
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/19
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal intNtCd As String) As Rec_MNutri
        Try
            Key = Nothing
            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            For Each objMsg As Rec_MNutri In InnerList
                If objMsg.NtCd = intNtCd Then
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
    ' Design On  :2007/12/03
    '        By  :CM)Kuwahara
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MNutri) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :�V�K���R�[�h���(�ėp)
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/04
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function NewRecord() As Object
        Dim objRecMNutri As Rec_MNutri

        objRecMNutri = New Rec_MNutri()
        Return objRecMNutri
    End Function

    '**********************************************************************
    ' Name       :�A�C�e���}��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/03
    '        By  :CM)Kuwahara
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MNutri)
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
    ' Design On  :2007/12/03
    '        By  :CM)Kuwahara
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
    ' Design On  :2007/12/03
    '        By  :CM)Kuwahara
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MNutri

        Dim objCol As New Col_MNutri
        '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
        For intIdx As Integer = 0 To Me.Count - 1
            ' Item���N���[�� �R�s�[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MNutri))
        Next
        '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^
        Return objCol

    End Function

End Class