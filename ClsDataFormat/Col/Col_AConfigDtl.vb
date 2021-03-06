Option Strict On
'**************************************************************************
' Name       :©®è`Ú×}X^RNV
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_AConfigDtl

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ACeQÆ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_AConfigDtl
        Get
            Try
                Return CType(InnerList.Item(index), Rec_AConfigDtl)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ACeÇÁ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_AConfigDtl) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :ACe}ü
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_AConfigDtl)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ACeí
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
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
    ' Comment    :ñöJoÆ·é 
    '
    '----------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_AConfigDtl

        Dim objCol As New Col_AConfigDtl
        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemðN[ Rs[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_AConfigDtl))
        Next
        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^
        Return objCol

    End Function

End Class