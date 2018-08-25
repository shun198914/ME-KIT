Option Strict On

<Serializable()> Public Class Col_GTblCst

    Inherits CollectionBase
    'Implements System.ICloneable







    '**********************************************************************
    ' Summary    : �A�C�e���Q��
    '----------------------------------------------------------------------
    ' Method     : Public Overloads ReadOnly Property Item()
    '
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 
    '        By  : 
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Default Public Overloads ReadOnly Property Item(ByVal index As Integer) As Rec_GTblCst
        Get
            Try
                Return CType(InnerList.Item(index), Rec_GTblCst)
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Summary    : �A�C�e���ǉ�
    '----------------------------------------------------------------------
    ' Method     : Public Function Add()
    '
    ' Parameter  : ByVal newValue As MSGUserInfo
    '
    ' Return     : Integer
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 
    '        By  : 
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Public Function Add(ByVal newValue As Rec_GTblCst) As Integer
        Try
            InnerList.Add(newValue)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Summary    : �A�C�e���}��
    '----------------------------------------------------------------------
    ' Method     : Public Sub Insert()
    '
    ' Parameter  : ByVal Index As Integer,ByVal newValue As MSGGetTblConst
    '
    ' Return     : 
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 
    '        By  : 
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub Insert(ByVal Index As Integer, ByVal newValue As Rec_GTblCst)
        Try
            InnerList.Insert(Index, newValue)
        Catch ex As Exception
        End Try
    End Sub

    '**********************************************************************
    ' Summary    : �A�C�e���ǉ�
    '----------------------------------------------------------------------
    ' Method     : Public Function Add()
    '
    ' Parameter  : ByVal intCode As Integer
    '              ByVal strName As String
    ' Return     : Integer
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Public Function Add(ByVal intCode As Integer, ByVal strName As String) As Integer
        Try
            Dim objMSGGetTblConst As Rec_GTblCst
            objMSGGetTblConst = New Rec_GTblCst
            objMSGGetTblConst.ConstCode = intCode
            objMSGGetTblConst.ConstName = strName
            InnerList.Add(objMSGGetTblConst)
            Return Count()
        Catch ex As Exception
            Return -1
        End Try
    End Function

    '**********************************************************************
    ' Summary    : �A�C�e���폜
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
    ' Design On  : 2004/05/19
    '        By  : FTC) S_SHIN
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
