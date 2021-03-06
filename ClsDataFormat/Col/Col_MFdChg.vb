Option Strict On
'**************************************************************************
' Name       :HνΚ­ΞΫHi}X^RNV
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2007/11/22
'        By  :FTC)Hasyashi
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MFdChg

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ACeQΖ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/11/22
    '        By  :FTC)Hasyashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MFdChg
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MFdChg)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    ''**********************************************************************
    '' Name       :ACeQΖidόζR[hj
    ''
    ''----------------------------------------------------------------------
    '' Comment    :
    ''
    ''----------------------------------------------------------------------
    '' Design On  :2008/04/07
    ''        By  :K.Taguchi
    ''----------------------------------------------------------------------
    '' Update On  :
    ''        By  :
    '' Comment    :
    ''**********************************************************************
    'Public Function Key(ByVal strTrndCd As String) As Rec_MFdChg
    '    Try
    '        Key = Nothing
    '        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
    '        For Each objMsg As Rec_MFdChg In InnerList
    '            If objMsg.TrndCd.Trim = strTrndCd.Trim Then
    '                Key = objMsg
    '                Exit For
    '            End If
    '        Next
    '        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^

    '        Return Key

    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function

    '**********************************************************************
    ' Name       :ACeΗΑ
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/11/22
    '        By  :FTC)Hasyashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MFdChg) As Integer
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
    ' Design On  :2007/11/22
    '        By  :FTC)Hasyashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MFdChg)
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
    ' Design On  :2007/11/22
    '        By  :FTC)Hasyashi
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
    ' Design On  :2007/11/22
    '        By  :FTC)Hasyashi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MFdChg

        Dim objCol As New Col_MFdChg
        '^PPPPPPPPPPPPPPPPPPPPPPPPPPPPP_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemπN[ Rs[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MFdChg))
        Next
        '_QQQQQQQQQQQQQQQQQQQQQQQQQQQQQ^
        Return objCol

    End Function

End Class
