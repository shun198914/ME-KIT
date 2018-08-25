Option Strict On
'**************************************************************************
' Name       :ÉRÉÅÉìÉgÉ}ÉXÉ^ÉRÉåÉNÉVÉáÉì
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2007/12/05
'        By  :FM)Sugiyama
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MCmt

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄéQè∆
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/05
    '        By  :FM)Sugiyama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MCmt
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MCmt)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄéQè∆ÅiÉRÉÅÉìÉgÉRÅ[ÉhÅj
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/04/25
    '        By  :K.Taguchi
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strCmtCd As String) As Rec_MCmt
        Try
            Key = Nothing
            'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
            For Each objMsg As Rec_MCmt In InnerList
                If objMsg.CmtCd.Trim = strCmtCd.Trim Then
                    Key = objMsg
                    Exit For
                End If
            Next
            'Å_ÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅ^

            Return Key

        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄí«â¡
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/05
    '        By  :FM)Sugiyama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MCmt) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄë}ì¸
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/05
    '        By  :FM)Sugiyama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MCmt)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄçÌèú
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/05
    '        By  :FM)Sugiyama
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
    ' Method     :ÉAÉCÉeÉÄÉNÉçÅ[Éì
    '
    '----------------------------------------------------------------------
    ' Comment    :îÒåˆäJÉÅÉìÉoÇ∆Ç∑ÇÈ 
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/12/05
    '        By  :FM)Sugiyama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MCmt

        Dim objCol As New Col_MCmt
        'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemÇÉNÉçÅ[Éì ÉRÉsÅ[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MCmt))
        Next
        'Å_ÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅ^
        Return objCol

    End Function

End Class