Option Strict On
'**************************************************************************
' Name       :食数集計画面設定マスタコレクション
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2016/08/19
'        By  :FTC)T.Takeuchi
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_SMenuCountDiff

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :アイテム参照
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2016/08/19
    '        By  :FTC)T.Takeuchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_SMenuCountDiff
        Get
            Try
                Return CType(InnerList.Item(index), Rec_SMenuCountDiff)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :アイテム追加
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2016/08/19
    '        By  :FTC)T.Takeuchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_SMenuCountDiff) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :アイテム挿入
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2016/08/19
    '        By  :FTC)T.Takeuchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_SMenuCountDiff)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Name       :アイテム削除
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2016/08/19
    '        By  :FTC)T.Takeuchi
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
    ' Method     :アイテムクローン
    '
    '----------------------------------------------------------------------
    ' Comment    :非公開メンバとする 
    '
    '----------------------------------------------------------------------
    ' Design On  :2016/08/19
    '        By  :FTC)T.Takeuchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_SMenuCountDiff

        Dim objCol As New Col_SMenuCountDiff
        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
        For intIdx As Integer = 0 To Me.Count - 1
            ' Itemをクローン コピー
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_SMenuCountDiff))
        Next
        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／
        Return objCol

    End Function

End Class