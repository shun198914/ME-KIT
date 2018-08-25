Option Strict On
'**************************************************************************
' Name       : 配信レコードコレクション（発注可能商品情報）
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  : 2008/06/13
'        By  : (FTC)K.Aoki
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_Dl_KHKS
    Inherits CollectionBase

    '**********************************************************************
    ' Summary    : アイテム参照
    '----------------------------------------------------------------------
    ' Method     : Public Overloads ReadOnly Property Item()
    '
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2008/06/13
    '        By  : (FTC)K.Aoki
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_Dl_KHKS
        Get
            Try
                Return CType(InnerList.Item(index), Rec_Dl_KHKS)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Summary    : アイテム追加
    '----------------------------------------------------------------------
    ' Method     : Public Function Add()
    '
    ' Parameter  : ByVal newValue As Rec_Dl_KHKS
    '
    ' Return     : Integer
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2008/06/13
    '        By  : (FTC)K.Aoki
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_Dl_KHKS) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Summary    : アイテム挿入
    '----------------------------------------------------------------------
    ' Method     : Public Sub Insert()
    '
    ' Parameter  : ByVal Index As Integer,ByVal newValue As Rec_Dl_KHKS
    '
    ' Return     : 
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2008/06/13
    '        By  : (FTC)K.Aoki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_Dl_KHKS)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Summary    : アイテム削除
    '----------------------------------------------------------------------
    ' Method     : Public Sub Remove()
    '
    ' Parameter  : ByVal Index As Integer
    '
    ' Return     : 
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2008/06/13
    '        By  : (FTC)K.Aoki
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
