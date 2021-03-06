Option Strict On
'**************************************************************************
' Name       :HiQ}X^1RNV
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/01/30
'        By  :y.hayashi
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MGpFd

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ACeQΖ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/30
    '        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MGpFd
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MGpFd)
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
    ' Design On  :2008/01/30
    '        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MGpFd) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :ACeQΖinζR[h+HiQR[hj
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/21
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strTikuCd As String _
                      , ByVal strFdGpCd As String) As Rec_MGpFd
        Try
            Key = Nothing
            '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
            For Each objMsg As Rec_MGpFd In InnerList
                If objMsg.TikuCd = strTikuCd _
                AndAlso objMsg.FdGpCd = strFdGpCd Then
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
    ' Name       :ACeQΖiHiQR[hj
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/21
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strFdGpCd As String) As Rec_MGpFd
        Try
            Key = Nothing
            '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
            For Each objMsg As Rec_MGpFd In InnerList
                If objMsg.FdGpCd = strFdGpCd Then
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
    ' Name       :ACe}ό
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/30
    '        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MGpFd)
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
    ' Design On  :2008/01/30
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
    ' Design On  :2008/01/30
    '        By  :y.hayashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MGpFd

        Dim objCol As New Col_MGpFd
        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemπN[ Rs[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MGpFd))
        Next
        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^
        Return objCol

    End Function

End Class