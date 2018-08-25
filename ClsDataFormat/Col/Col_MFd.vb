Option Strict On
'**************************************************************************
' Name       :êHïiÉ}ÉXÉ^ÉRÉåÉNÉVÉáÉì
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2008/01/15
'        By  :FMS)Y.Maki
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MFd

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄéQè∆
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MFd
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MFd)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄéQè∆ÅiêHïiÉ}ÉXÉ^Åj
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/05/16
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Key(ByVal strFdCd As String) As Rec_MFd
        Try
            Key = Nothing
            'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
            For Each objMsg As Rec_MFd In InnerList
                If objMsg.FdCd.Trim = strFdCd.Trim Then
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
    ' Name       :êVãKÉåÉRÅ[ÉhèÓïÒ(îƒóp)
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/02/02
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function NewRecord() As Object
        Dim objRecMFd As Rec_MFd

        objRecMFd = New Rec_MFd()
        Return objRecMFd
    End Function

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄí«â¡
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MFd) As Integer
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MFd)
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
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
    ' Design On  :2008/01/15
    '        By  :FMS)Y.Maki
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MFd

        Dim objCol As New Col_MFd
        'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemÇÉNÉçÅ[Éì ÉRÉsÅ[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MFd))
        Next
        'Å_ÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅ^
        Return objCol

    End Function

End Class