Option Strict On
'**************************************************************************
' Name       :ΏHi}X^RNV
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MCokFd

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ACeQΖ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MCokFd
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MCokFd)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ACeΗΑ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MCokFd) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :ACe}ό
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MCokFd)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ACeν
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
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
    ' Method     :ACeN[
    '
    '----------------------------------------------------------------------
    ' Comment    :ρφJoΖ·ι 
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/16
'        By  :aaa
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MCokFd

        Dim objCol As New Col_MCokFd
        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemπN[ Rs[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MCokFd))
        Next
        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^
        Return objCol

    End Function

End Class