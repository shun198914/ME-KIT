Option Strict On
'**************************************************************************
' Name       :Ε½HiQ}X^RNV
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/01/19
'        By  :y.hayashi
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MGpFdMx

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ACeQΖ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
' Design On  :2008/01/19
'        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MGpFdMx
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MGpFdMx)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ACeQΖiΕ½HiQR[hj
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/04/30
    '        By  :K.Taguchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strGpFdMxCd As String) As Rec_MGpFdMx
        Try
            Key = Nothing
            '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
            For Each objMsg As Rec_MGpFdMx In InnerList
                If objMsg.GpFdMxCd.Trim = strGpFdMxCd.Trim Then
                    Key = objMsg
                    Exit For
                End If
            Next
            '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^

            Return Key

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '**********************************************************************
    ' Name       :ACeΗΑ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/19
    '        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MGpFdMx) As Integer
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
' Design On  :2008/01/19
'        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MGpFdMx)
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
' Design On  :2008/01/19
'        By  :y.hayashi
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
' Design On  :2008/01/19
'        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MGpFdMx

        Dim objCol As New Col_MGpFdMx
        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemπN[ Rs[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MGpFdMx))
        Next
        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^
        Return objCol

    End Function

End Class