Option Strict On
'**************************************************************************
' Name       :êHéÌÉRÉÅÉìÉgÉ}ÉXÉ^Åiå¬êlÅjÉRÉåÉNÉVÉáÉì
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :2010/11/12
'        By  :FTC) S.Shin
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
<Serializable()> Public Class Col_MMealCmtP

    Inherits CollectionBase
    Implements System.ICloneable

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄéQè∆
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2010/11/12
    '        By  :FTC) S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_MMealCmtP
        Get
            Try
                Return CType(InnerList.Item(index), Rec_MMealCmtP)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property
    Default Public Overloads ReadOnly Property Item(ByVal CmtCd As String) As Col_MMealCmtP
        Get
            Try
                Dim Col As New Col_MMealCmtP

                'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
                For Each Rec As Rec_MMealCmtP In InnerList
                    If Rec.CmtCd.Trim.ToUpper = CmtCd.Trim.ToUpper Then
                        Col.Add(Rec)
                    End If
                Next
                'Å_ÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅ^

                Return Col

            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :ÉAÉCÉeÉÄí«â¡
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2010/11/12
    '        By  :FTC) S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_MMealCmtP) As Integer
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
' Design On  :2010/11/12
'        By  :FTC) S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_MMealCmtP)
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
' Design On  :2010/11/12
'        By  :FTC) S.Shin
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
' Design On  :2010/11/12
'        By  :FTC) S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CloneR() As Object Implements System.ICloneable.Clone
        Return Me.MemberwiseClone()
    End Function
    Public Function Clone() As Col_MMealCmtP

        Dim objCol As New Col_MMealCmtP
        'Å^ÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅPÅ_
        For intIdx As Integer = 0 To Me.Count - 1
            ' ItemÇÉNÉçÅ[Éì ÉRÉsÅ[
            objCol.Add(DirectCast(Me.Item(intIdx).Clone(), Rec_MMealCmtP))
        Next
        'Å_ÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅQÅ^
        Return objCol

    End Function

End Class