Option Strict On
'**************************************************************************
' Name       :ClsCom ���ʃ��C�u����
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :AST)M.Yoshida
'        By  :2007/12/01
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
'**************************************************************************
' Imports
'**************************************************************************
Imports ClsDataFormat
'**************************************************************************
' Class
'**************************************************************************
Public Class ClsCom

    '**********************************************************************
    ' �N���X�ŗL�����o�[
    '**********************************************************************
    Private Shared m_recXSys As Rec_XSys
    '**********************************************************************
    '�C�x���g
    '**********************************************************************
    Public Shared Event CmnLogOut(ByVal ex As Exception _
                                , ByVal enmLogType As ClsCst.enmLogType _
                                , ByVal strText As String _
                                , ByVal blnDebugLogOutOnly As Boolean)

    '**********************************************************************
    ' Name       :���O/��O�o��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2010/11/28
    '        By  :ATST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Shared Sub LogOut(ByVal enmLogType As ClsCst.enmLogType _
                    , ByRef strText As String _
                    , Optional ByVal blnDebugLogOutOnly As Boolean = False)
        RaiseEvent CmnLogOut(Nothing, enmLogType, strText, blnDebugLogOutOnly)
    End Sub
    Private Shared Sub LogOut(ByVal ex As Exception)
        RaiseEvent CmnLogOut(ex, ClsCommon.ClsCst.enmLogType.Error, Nothing, False)
    End Sub

    '**********************************************************************
    ' Name       :XML�t�@�C���f�[�^�擾  
    '
    '----------------------------------------------------------------------
    ' Comment    :strFileName  'XML�t�@�C����
    '            :strSection   '�Z�N�V������
    '            :strKey �@�@  '�L�[��
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function GetXmlData(ByRef strFileName As String, ByRef strSection As String, ByRef strKey As String) As String

        Try
            '�߂�l�̏�����
            Dim strReturn As String = ""

            ''XML�t�@�C����ǂݍ��ށB
            Dim objXmlDocument As New System.Xml.XmlDocument
            objXmlDocument.Load(strFileName)

            '�w��Z�N�V�����̃m�[�h���X�g���擾
            Dim objNodeList As System.Xml.XmlNodeList = objXmlDocument.GetElementsByTagName(strSection)
            If Not objNodeList Is Nothing Then
                '�����Z�N�V���������������ꍇ�擪�̃m�[�h���g�p����
                Dim childNodeList As System.Xml.XmlNodeList = objNodeList.Item(0).ChildNodes
                If Not childNodeList Is Nothing Then
                    '�Z�N�V�����̒��̃m�[�h�i�w��L�[�j������
                    For Each objNode As System.Xml.XmlNode In childNodeList
                        If objNode.Name = strKey Then
                            '�ݒ�l�擾���߂�l�Ƃ��ĕԂ�
                            strReturn = objNode.InnerText
                            Exit For
                        End If
                    Next
                End If
            End If

            GetXmlData = strReturn

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�V�X�e���ݒ���
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2007/10/29
    '        By  :AST)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Property SysInf() As Rec_XSys
        Get
            Return m_recXSys
        End Get
        Set(ByVal Value As Rec_XSys)
            m_recXSys = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :�A�v���P�[�V�����p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By  : 
    ' Comment    : 
    '**********************************************************************
    Public Shared ReadOnly Property GetAppPath() As String
        Get
            Try
                GetAppPath = System.AppDomain.CurrentDomain.BaseDirectory
                Do Until GetAppPath = "\"
                    If System.IO.File.Exists(GetAppPath & ClsCst.FileName.ConfigClient) _
                    Or System.IO.File.Exists(GetAppPath & ClsCst.FileName.ConfigServer) Then
                        Exit Do
                    End If
                    GetAppPath = String.Join("\"c, GetAppPath.Split("\"c), 0, GetAppPath.Split("\"c).Length - 2) & "\"c
                Loop
                If GetAppPath = "\" Then
                    ''OutPutLog("�̏�ʂ�" & ClsCommon.ClsCst.FldrPath.System & "�t�H���_��������܂���ł����B")
                End If
                '�V�X�e���p�X��Ԃ��܂��B
                Return GetAppPath

            Catch ex As Exception
                Throw
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :�C���[�W�p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On�@: 
    '        By�@: 
    ' Comment  �@: 
    '**********************************************************************
    Public Shared ReadOnly Property GetImagePath() As String
        Get
            '�摜�p�X��Ԃ��܂��B
            Return GetAppPath & ClsCst.FldrName.Image & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :�T�E���h�p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On�@: 
    '        By�@: 
    ' Comment  �@: 
    '**********************************************************************
    Public Shared ReadOnly Property GetSoundPath() As String
        Get
            '�摜�p�X��Ԃ��܂��B
            Return GetAppPath & ClsCst.FldrName.Sound & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :���O�t�@�C���p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On : 
    '        By : 
    ' Comment   : 
    '**********************************************************************
    Public Shared ReadOnly Property GetLogPath() As String
        Get
            '���O�t�@�C���p�X��Ԃ��܂��B
            Return GetAppPath & ClsCst.FldrName.Log & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :SQL�X�N���v�g�t�@�C���p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/05/09
    '----------------------------------------------------------------------
    ' Update On : 
    '        By : 
    ' Comment   : 
    '**********************************************************************
    Public Shared ReadOnly Property GetSqlScriptPath() As String
        Get
            'SQL�X�N���v�g�t�@�C���p�X��Ԃ��܂��B
            Return GetAppPath & ClsCst.FldrName.SqlScript & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :�w���v�t�@�C���p�X
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '�@�@�@�@�@�@�@
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By�@: 
    ' Comment    : 
    '**********************************************************************
    Public Shared ReadOnly Property GetHelpPath() As String
        Get
            '�摜�p�X��Ԃ��܂��B
            Return GetAppPath & ClsCst.FldrName.Help & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :����������p���ŋ���
    '
    '----------------------------------------------------------------------
    ' Comment    :strText  ������
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvQ(ByRef strText As String) As String
        Try
            '���X�̕����񂪈��p�����܂ޏꍇ�̕ϊ����K�v
            Return "'" & strText.Replace("'", "''") & "'"
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :�_�u���N�H�[�g�ҏW�i�I�[�g�j
    '
    '----------------------------------------------------------------------
    ' Comment    :�����^�A�L�����N�^�^�A���t�^�́A�_�u���N�H�[�g�ҏW��A�O��"�t��
    '�@�@�@�@�@  �@ ����ȊO�͒P����String�ϊ����ĕԂ��܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvSQL(ByVal objMember As Object) As String
        Try
            Select Case True
                Case TypeOf objMember Is String _
                   , TypeOf objMember Is Char _
                   , TypeOf objMember Is DateTime
                    '���X�̕����񂪈��p�����܂ޏꍇ�̕ϊ����K�v
                    Return "'" & CType(objMember, String).Replace("'", "''") & "'"
                Case Else
                    '���̂܂ܕ����񉻂��ĕԂ�
                    Return CType(objMember, String)
            End Select

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :����������s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoin(ByRef strBase As String, ByRef strAddText As String)
        Try
            'CrLf�ōs���������܂��B
            strBase = strBase & ControlChars.CrLf & strAddText
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :����������s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :FMS)Y.Maki
    '        By  :2007/12/20
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoinNormal(ByRef strBase As System.Text.StringBuilder, ByVal strAddText As String)
        Try
            'CrLf�ōs���������܂��B
            strBase.Append(ControlChars.CrLf & strAddText)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :��������J���}�����t�����s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :FMS)Y.Maki
    '        By  :2007/12/20
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoinItem(ByRef strBase As System.Text.StringBuilder, ByVal strAddText As String, ByRef blnItemSecond As Boolean)
        Try
            Dim strJoint As String = String.Empty

            If strAddText.Length <= 0 Then
                Return
            End If
            If blnItemSecond = False Then
                blnItemSecond = True
            Else
                strJoint = "    , "
            End If

            'CrLf�ōs���������܂��B
            strBase.Append(ControlChars.CrLf & strJoint & strAddText)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :�������WHERE����AND�����t�����s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :FMS)Y.Maki
    '        By  :2007/12/20
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoinWhere(ByRef strBase As System.Text.StringBuilder, ByVal strItemName As String, ByVal objItemValue As Object, ByRef blnWhereSecond As Boolean)
        Try
            Dim strJoint As String = String.Empty

            If (objItemValue Is Nothing) OrElse (CType(objItemValue, String).Length <= 0) Then
                Return
            End If

            If blnWhereSecond = False Then
                strJoint = "     WHERE "
                blnWhereSecond = True
            Else
                strJoint = "     AND "
            End If
            strJoint = strJoint & strItemName & " = " & CnvSQL(objItemValue)

            'CrLf�ōs���������܂��B
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :�������ORDER BY���̓J���}�����t�����s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :FMS)Y.Maki
    '        By  :2007/12/20
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoinOrder(ByRef strBase As System.Text.StringBuilder, ByVal strItemName As String, ByRef blnOrderSecond As Boolean, Optional ByVal blnOrderDescending As Boolean = False)
        Try
            Dim strJoint As String = String.Empty

            If strItemName.Length <= 0 Then
                Return
            End If
            If blnOrderSecond = False Then
                blnOrderSecond = True
                strJoint = "    ORDER BY "
            Else
                strJoint = "    , "
            End If
            strJoint = strJoint & strItemName
            If blnOrderDescending = True Then
                strJoint = strJoint & " DESC"
            End If

            'CrLf�ōs���������܂��B
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :�������SET���̓J���}�����t�����s�ŘA�����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub strRowJoinUpdate(ByRef strBase As System.Text.StringBuilder, ByVal strItemName As String, ByVal objItemValue As Object, ByRef blnItemSecond As Boolean, ByVal blnItemPersonal As Boolean, ByVal blnItemOrder As Boolean)
        Try
            Dim strJoint As String = String.Empty

            If (blnItemPersonal) AndAlso (blnItemOrder = False) Then
                Return
            End If
            If (objItemValue Is Nothing) OrElse (CType(objItemValue, String).Length <= 0) Then
                objItemValue = Nothing
            End If

            If blnItemSecond = False Then
                strJoint = " SET "
                blnItemSecond = True
            Else
                strJoint = ", "
            End If
            strJoint = strJoint & strItemName & " = "
            If strItemName = "UpDtTime" Then
                strJoint = strJoint & "GetDate()"
            ElseIf objItemValue Is Nothing Then
                strJoint = strJoint & "null"
            Else
                strJoint = strJoint & CnvSQL(objItemValue)
            End If

            'CrLf�ōs���������܂��B
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :�t�@�C���^�t�H���_�p�X���̑�֕�����ϊ�
    '
    '----------------------------------------------------------------------
    ' Comment    :%AppRoot% ��FrmMenu.Exe�̂���t�H���_�p�X�ɂ��܂��B
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvPath(ByRef strPath As String) As String
        Try
            CnvPath = ""
            '�ȉ��ɏے������񕪂̃��v���[�X���������ĉ������B
            CnvPath = Replace(strPath, ClsCst.FldrName.AppRoot, GetAppPath, 1, -1, CompareMethod.Text)
            'CnvPath = strPath.Replace(ClsCst.FldrName.AppRoot, GetAppPath)

            Return CnvPath

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :������̈Í���
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function EncryptByRC2(ByRef pKeyWord As String) As String
        Try
            EncryptByRC2 = Nothing

            Dim objCryptoServiceProvider As New System.Security.Cryptography.RC2CryptoServiceProvider '�Í��������{����I�u�W�F�N�g
            Dim bytInPutWords As Byte() = System.Text.Encoding.UTF8.GetBytes(pKeyWord)                '�Í����Ώۃo�C�g��i�p�����^���o�C�g��ɃG���R�[�h�j
            Dim bytOutPutWords As Byte()                                                              '�Í�����̃o�C�g��
            Dim objOutPutSteam As New System.IO.MemoryStream                                          '�Í����f�[�^���������ރI�u�W�F�N�g

            '�Í������ݒ�
            objCryptoServiceProvider.Key = System.Text.Encoding.UTF8.GetBytes("YkAeMnAjSiH$I\TA")

            '�������x�N�^�ݒ�
            objCryptoServiceProvider.IV = System.Text.Encoding.UTF8.GetBytes("AjiDAS-6")

            'CryptoStream�I�u�W�F�N�g�̍쐬
            Dim objTransform As System.Security.Cryptography.ICryptoTransform = objCryptoServiceProvider.CreateEncryptor
            Dim objCryptoStream As New System.Security.Cryptography.CryptoStream(objOutPutSteam, objTransform, System.Security.Cryptography.CryptoStreamMode.Write)

            '�Í�������Stream�ɏo��
            objCryptoStream.Write(bytInPutWords, 0, bytInPutWords.Length)
            objCryptoStream.FlushFinalBlock()

            'Stream����Í����f�[�^���擾
            bytOutPutWords = objOutPutSteam.ToArray

            '�I�u�W�F�N�g�����
            objCryptoStream.Close()
            objOutPutSteam.Close()

            '���A�l��ݒ肷��
            EncryptByRC2 = System.Convert.ToBase64String(bytOutPutWords)
        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :������̕�����
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function DecryptByRC2(ByRef pKeyWord As String) As String
        Try
            DecryptByRC2 = Nothing

            Dim objCryptoServiceProvider As New System.Security.Cryptography.RC2CryptoServiceProvider '�����������{����I�u�W�F�N�g
            Dim bytInPutWords As Byte() = System.Convert.FromBase64String(pKeyWord)                   '�������Ώۃo�C�g��i�p�����^��BASE64�Ńf�R�[�h�j
            Dim strOutPutWords As String                                                              '��������̕�����
            Dim objInPutSteam As New System.IO.MemoryStream(bytInPutWords)                            '�������f�[�^��ǂݍ��ރI�u�W�F�N�g

            '�Í������ݒ�
            objCryptoServiceProvider.Key = System.Text.Encoding.UTF8.GetBytes("YkAeMnAjSiH$I\TA")

            '�������x�N�^�ݒ�
            objCryptoServiceProvider.IV = System.Text.Encoding.UTF8.GetBytes("AjiDAS-6")

            'CryptoStream�I�u�W�F�N�g�̍쐬
            Dim objTransform As System.Security.Cryptography.ICryptoTransform = objCryptoServiceProvider.CreateDecryptor
            Dim objCryptoStream As New System.Security.Cryptography.CryptoStream(objInPutSteam, objTransform, System.Security.Cryptography.CryptoStreamMode.Read)

            '����������Stream�������
            Dim objStreamReader As New System.IO.StreamReader(objCryptoStream, System.Text.Encoding.UTF8)
            strOutPutWords = objStreamReader.ReadToEnd

            '�I�u�W�F�N�g�����
            objStreamReader.Close()
            objCryptoStream.Close()
            objInPutSteam.Close()

            '���A�l��ݒ肷��
            DecryptByRC2 = strOutPutWords
        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�C���[�W���X�g�ւ̉摜�ǂݍ���
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function SetImageList(ByRef iltCtl As System.Windows.Forms.ImageList _
                                      , ByRef strFileFormat As String) As Integer
        Try
            '�߂�l�̏�����
            SetImageList = -1

            Dim strImagePath As String = GetImagePath

            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            For Each strLogFileName As String In System.IO.Directory.GetFiles(strImagePath, strFileFormat)
                Dim objImg As System.Drawing.Image = System.Drawing.Image.FromFile(strLogFileName)
                iltCtl.ImageSize = New System.Drawing.Size(objImg.Width, objImg.Height)
                iltCtl.Images.Add(System.IO.Path.GetFileName(strLogFileName), objImg)
                '�߂�l���Ō�ɒǉ������A�C�e���̃C���f�b�N�X�l�ɂ��܂��B
                SetImageList = iltCtl.Images.Count - 1
            Next
            '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^


        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :WAVE�t�@�C���̍Đ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Sub PlaySound(ByVal waveFile As String)
        Try
            '�ǂݍ���
            Dim player As New System.Media.SoundPlayer(GetSoundPath & waveFile)

            '�񓯊��Đ�����
            player.Play()

            'Throw New DivideByZeroException

            'Beep()


            '���̂悤�ɂ���ƁA���[�v�Đ������
            'player.PlayLooping()

            '���̂悤�ɂ���ƁA�Ō�܂ōĐ����I����܂őҋ@����
            'player.PlaySync()
        Catch ex As Exception
            Throw
        End Try
    End Sub


    '**********************************************************************
    ' Name       :�I��H�敪���I��H��
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/10
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvSelEtDivToName(ByRef intEatTimeDiv As Integer) As String
        Try
            CnvSelEtDivToName = ""
            If intEatTimeDiv < 1 OrElse intEatTimeDiv > 26 Then Exit Try

            Return ("ABCDEFGHIJKLMNOPQRSTUVWXYZ").Substring(intEatTimeDiv - 1, 1)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�I��H�����I��H�敪
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/10
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvSelEtNameToDiv(ByVal strSelEtName As String) As Integer
        Try
            If strSelEtName.Trim = String.Empty Then
                Return -1
            Else
                Return ("ABCDEFGHIJKLMNOPQRSTUVWXYZ").IndexOf(strSelEtName.ToUpper.Trim) + 1
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����ԋ敪���H�����Ԗ�(���͗���)
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvEatTimeDivToName(ByRef enumEatTimeDiv As ClsCst.enmEatTimeDiv _
                                    , Optional ByVal blnShortName As Boolean = False) As String
        Try
            CnvEatTimeDivToName = ""

            If blnShortName = False Then
                Select Case enumEatTimeDiv
                    Case ClsCst.enmEatTimeDiv.MorningDiv : Return ClsCst.EatTime.MorningName
                    Case ClsCst.enmEatTimeDiv.AM10Div : Return ClsCst.EatTime.AM10DivName
                    Case ClsCst.enmEatTimeDiv.LunchDiv : Return ClsCst.EatTime.LunchName
                    Case ClsCst.enmEatTimeDiv.PM15Div : Return ClsCst.EatTime.PM15Name
                    Case ClsCst.enmEatTimeDiv.EveningDiv : Return ClsCst.EatTime.EveningName
                    Case ClsCst.enmEatTimeDiv.NightDiv : Return ClsCst.EatTime.NightName
                End Select
            Else
                Select Case enumEatTimeDiv
                    Case ClsCst.enmEatTimeDiv.MorningDiv : Return ClsCst.EatTime.MorningNameShort
                    Case ClsCst.enmEatTimeDiv.AM10Div : Return ClsCst.EatTime.AM10DivNameShort
                    Case ClsCst.enmEatTimeDiv.LunchDiv : Return ClsCst.EatTime.LunchNameShort
                    Case ClsCst.enmEatTimeDiv.PM15Div : Return ClsCst.EatTime.PM15NameShort
                    Case ClsCst.enmEatTimeDiv.EveningDiv : Return ClsCst.EatTime.EveningNameShort
                    Case ClsCst.enmEatTimeDiv.NightDiv : Return ClsCst.EatTime.NightNameShort
                End Select
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����Ԗ�(���͗���)���H�����ԋ敪
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvEatTimeNameToDiv(ByRef EatTimeName As String) As ClsCst.enmEatTimeDiv
        Try
            CnvEatTimeNameToDiv = ClsCst.enmEatTimeDiv.None

            Select Case EatTimeName
                Case ClsCst.EatTime.MorningName, ClsCst.EatTime.MorningNameShort : Return ClsCst.enmEatTimeDiv.MorningDiv
                Case ClsCst.EatTime.AM10DivName, ClsCst.EatTime.AM10DivNameShort : Return ClsCst.enmEatTimeDiv.AM10Div
                Case ClsCst.EatTime.LunchName, ClsCst.EatTime.LunchNameShort : Return ClsCst.enmEatTimeDiv.LunchDiv
                Case ClsCst.EatTime.PM15Name, ClsCst.EatTime.PM15NameShort : Return ClsCst.enmEatTimeDiv.PM15Div
                Case ClsCst.EatTime.EveningName, ClsCst.EatTime.EveningNameShort : Return ClsCst.enmEatTimeDiv.EveningDiv
                Case ClsCst.EatTime.NightName, ClsCst.EatTime.NightNameShort : Return ClsCst.enmEatTimeDiv.NightDiv
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :��������e�`�F�b�N
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function StringCheck(ByVal ChkInf As Rec_XChkInf _
                            , Optional ByVal blnTrimAndPading As Boolean = False) As Rec_XChkInf

        Try
            '�G���[�t���O�̏�����
            ChkInf.Err = False
            ChkInf.LenBetweenErr = False
            ChkInf.LenTextMaxErr = False
            ChkInf.LenTextMinErr = False
            ChkInf.EmptyErr = False
            ChkInf.WideCharErr = False
            ChkInf.CrLfTabErr = False
            ChkInf.FilCtlCharErr = False
            ChkInf.NumericErr = False
            ChkInf.DecimalErr = False
            ChkInf.NumBetweenErr = False
            ChkInf.NumericMaxErr = False
            ChkInf.NumericMinErr = False
            ChkInf.DecLenErr = False

            Dim intstrLen As Integer = 0
            Dim strText As String = ChkInf.Text

            'ShiftJis���[�h�̏ꍇ
            Dim objEnc As System.Text.Encoding = Nothing
            If ChkInf.SJisMode = True Then
                '�G���R�[�h�I�u�W�F�N�g��p�ӂ��܂��B
                objEnc = System.Text.Encoding.GetEncoding("shift-jis")
                '�������Sjis�ϊ����������ɂ��܂��B
                strText = objEnc.GetString(objEnc.GetBytes(strText))
                '������̃o�C�g�����擾���܂��B
                intstrLen = objEnc.GetByteCount(strText)
            End If

            If blnTrimAndPading = True Then
                '�g���������w�莞�̓g������̕�����ɂ��܂��B
                Select Case ChkInf.TrimMode
                    Case Rec_XChkInf.enmTrim.Left : strText = strText.TrimStart
                    Case Rec_XChkInf.enmTrim.Right : strText = strText.TrimEnd
                    Case Rec_XChkInf.enmTrim.LeftRight : strText = strText.Trim
                End Select

                '�����ߎw�莞�͌����ߌ�̕�����ɂ��܂��B
                If ChkInf.TextMaxLen <> 0 AndAlso ChkInf.PadNarrowChar <> "" Then
                    If ChkInf.SJisMode = True Then
                        'ShiftJis���[�h�̏ꍇ
                        Select Case ChkInf.PadMode
                            Case Rec_XChkInf.enmPadMode.Left : strText = New String(ChkInf.PadNarrowChar, ChkInf.TextMaxLen - intstrLen) & strText
                            Case Rec_XChkInf.enmPadMode.Right : strText = strText & New String(ChkInf.PadNarrowChar, ChkInf.TextMaxLen - intstrLen)
                        End Select
                    Else
                        '�����łȂ��ꍇ
                        Select Case ChkInf.PadMode
                            Case Rec_XChkInf.enmPadMode.Left : strText = strText.PadLeft(ChkInf.TextMaxLen, ChkInf.PadNarrowChar)
                            Case Rec_XChkInf.enmPadMode.Right : strText = strText.PadRight(ChkInf.TextMaxLen, ChkInf.PadNarrowChar)
                        End Select
                    End If
                End If
            End If

            '������̒��������߂Ď擾���܂��B
            If ChkInf.SJisMode = True Then
                'ShiftJis���[�h�̏ꍇ�̓o�C�g�����擾���܂��B
                intstrLen = objEnc.GetByteCount(strText)
            Else
                '����ȊO�͏����ȕ��������擾���܂��B
                intstrLen = strText.Length
            End If

            '�G���R�[�h�I�u�W�F�N�g��j��
            objEnc = Nothing

            '-----------------------------------------------------------------
            '�e��`�F�b�N
            '-----------------------------------------------------------------

            '�K�{���͂Ŗ����͂Ȃ�G���[
            If ChkInf.NotEmpty = True AndAlso strText = String.Empty Then
                ChkInf.EmptyErr = True
            End If

            '�������̍ő包�w�肪����ꍇ�A�w�蕶�����𒴂�����G���[
            If ChkInf.TextMaxLen <> 0 AndAlso ChkInf.TextMaxLen < intstrLen Then
                ChkInf.LenTextMaxErr = True
            End If

            '�������̍ŏ����w�肪����ꍇ�A�w�蕶�����ȉ��Ȃ�G���[
            If ChkInf.TextMinLen <> 0 AndAlso ChkInf.TextMinLen > intstrLen Then
                ChkInf.LenTextMinErr = True
            End If

            '�������̍ő�or�ŏ����w�肪����ꍇ
            If ChkInf.LenTextMaxErr = True OrElse ChkInf.LenTextMinErr = True Then
                ChkInf.LenBetweenErr = True
            End If

            '�S�p�����֎~�ŁA���p�����ȊO�̕������܂܂�Ă���ꍇ�̓G���[
            If ChkInf.WideCharSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "^[ -~�-�\n\f\r\t]*$") = False Then
                ChkInf.WideCharErr = True
            End If

            '���sTab�����֎~�ŁA���sTab�������܂܂�Ă���ꍇ�̓G���[
            If ChkInf.CrLfTabSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "[\n\f\r\t]") = True Then
                ChkInf.CrLfTabErr = True
            End If

            '["][,]�̓��͋֎~�ŁA[']["][,]�������܂܂�Ă���ꍇ�̓G���[(��NumericMode=False�̎��̂ݗL��)
            If ChkInf.NumericMode = False Then
                If ChkInf.FilCtlCharSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "["",]") = True Then
                    ChkInf.FilCtlCharErr = True
                End If
            End If

            '���l�����w��̏ꍇ
            If ChkInf.DecimalOnly = True AndAlso strText.Trim Like New String("#"c, strText.Trim.Length) = False Then
                ChkInf.DecimalErr = True
            End If

            '���l�w��̏ꍇ
            If ChkInf.NumericMode = True AndAlso strText <> String.Empty Then
                '���l�łȂ��Ȃ�G���[
                Dim strChkNum As String = strText
                If strChkNum.Substring(0, 1) = "." Then strChkNum = "0." & strChkNum.Remove(0, 1)
                If strChkNum.Length > 1 AndAlso strChkNum.Substring(0, 2) = "-." Then strChkNum = "-0." & strChkNum.Remove(0, 2)
                If strChkNum = "-" Then strChkNum = "-0"

                If IsNumeric(strChkNum) = False _
                OrElse strChkNum.Trim.Replace(".", "0").Replace("-", "0") Like New String("#"c, strChkNum.Trim.Length) = False _
                OrElse (ChkInf.DecMaxLen = 0 AndAlso strChkNum.IndexOf("."c) <> -1) _
                OrElse ((ChkInf.NumMinValue >= 0 AndAlso ChkInf.NumMaxValue >= 0) AndAlso strChkNum.IndexOf("-"c) <> -1) Then
                    ChkInf.NumericErr = True
                Else
                    If strChkNum <> "."c Then
                        '���l�̍ő�^�ŏ��l�w�肪����ꍇ
                        If ChkInf.NumMaxValue <> 0D OrElse ChkInf.NumMinValue <> 0D Then
                            '�ŏ��l�ɖ����Ȃ��Ȃ�G���[
                            If ChkInf.NumMinValue > CType(strChkNum, Decimal) Then
                                ChkInf.NumericMinErr = True
                            End If
                            '�ő�l�𒴂�����G���[
                            If ChkInf.NumMaxValue < CType(strChkNum, Decimal) Then
                                ChkInf.NumericMaxErr = True
                            End If
                            '�ő�^�ŏ��l�𒴂��Ă���Ȃ�G���[
                            If ChkInf.NumericMinErr = True OrElse ChkInf.NumericMaxErr = True Then
                                ChkInf.NumBetweenErr = True
                            End If
                        End If
                        '�����_�̎w�茅
                        If ChkInf.DecMaxLen = 0 AndAlso strChkNum.IndexOf("."c) <> -1 Then
                            ChkInf.DecLenErr = True

                        ElseIf ChkInf.DecMaxLen > 0 AndAlso strChkNum.Split("."c).Length = 2 AndAlso strChkNum.Split("."c)(1).Length > ChkInf.DecMaxLen Then
                            ChkInf.DecLenErr = True
                        End If
                    End If
                End If
            End If

            '�e�L�X�g�{�b�N�X�̕ҏW�㕶����S�̂�Ԃ��B
            ChkInf.Text = strText

            '�ЂƂł��G���[������΁ATrue��ݒ肵�܂��B
            ChkInf.Err = (ChkInf.LenBetweenErr _
                       Or ChkInf.EmptyErr _
                       Or ChkInf.WideCharErr _
                       Or ChkInf.CrLfTabErr _
                       Or ChkInf.FilCtlCharErr _
                       Or ChkInf.NumericErr _
                       Or ChkInf.DecimalErr _
                       Or ChkInf.NumBetweenErr _
                       Or ChkInf.DecLenErr)

            Return ChkInf

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :Integer�͈̔̓`�F�b�N
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function ChkBetWeen(ByRef intValue As Integer _
                                    , ByRef intMinValue As Integer _
                                    , ByRef intMaxValue As Integer) As Boolean
        Try
            Return (intValue >= intMinValue AndAlso intValue <= intMaxValue)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :String�͈̔̓`�F�b�N
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/02/13
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function ChkBetWeen(ByRef strValue As String _
                                    , ByRef strMinValue As String _
                                    , ByRef strMaxValue As String) As Boolean
        Try
            Return (strValue >= strMinValue AndAlso strValue <= strMaxValue)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :Double�͈̔̓`�F�b�N
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function ChkBetWeen(ByRef dblValue As Double _
                                    , ByRef dblMinValue As Double _
                                    , ByRef dblMaxValue As Double) As Boolean
        Try
            Return (dblValue >= dblMinValue AndAlso dblValue <= dblMaxValue)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :���t�͈̔̓`�F�b�N
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function ChkBetWeen(ByRef dteValue As Date _
                                    , ByRef dteMinValue As Date _
                                    , ByRef dteMaxValue As Date) As Boolean
        Try
            Return (dteValue >= dteMinValue AndAlso dteValue <= dteMaxValue)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�����E����E���q�𕪐�������ɕϊ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionString(ByVal IntNo As Integer _
                                           , ByVal Bunshi As Integer _
                                           , ByVal Bunbo As Integer) As String
        Try
            '�m��̏ꍇ�̓A�b�v�_�E���R�����g�[�����番����������쐬
            Dim strText As String = String.Empty
            If IntNo <> 0 Then
                strText &= IntNo.ToString
            End If
            If Bunbo <> 0 Then
                If strText <> String.Empty Then strText &= "��"
                strText &= Bunshi.ToString & "/" & Bunbo.ToString
            End If
            Return strText.Trim

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :���������񂩂琮���E���q�E�����Integeer()���擾
    '
    '----------------------------------------------------------------------
    ' Comment    :�Ԃ����z��v�f��0:�����A1:���q�A2:����ł��B
    '             strUseCr��String.Empty�Ȃ炻�ꂼ��̒l��0�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionString(ByVal Value As Double) As String

        CnvFractionString = String.Empty
        Dim intNo As Long = CType(Math.Floor(Value), Integer) '������
        If intNo <> 0 Then Value = Value - intNo

        Dim baseValue As Long = 100 '�������̍ő包������0��ǉ�

        Dim Second As Long = baseValue
        Dim lngValue As Long = CType(Value * Second, Long)
        Dim lngGcm As Long = 0

        '���[�N���b�h�̌ݏ��@�ɂ�鐸��
        '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
        Do While Second > 0
            lngGcm = Second
            Second = lngValue Mod Second
            lngValue = lngGcm
        Loop
        '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^

        If intNo <> 0 Then
            '������������ꍇ
            CnvFractionString = intNo.ToString
        End If
        If Value <> 0 Then
            '������������ꍇ
            If CnvFractionString <> String.Empty Then CnvFractionString &= "��"
            CnvFractionString &= ((Value * baseValue) / lngGcm).ToString & "/" & (baseValue / lngGcm).ToString
        End If
        Return CnvFractionString

    End Function

    '**********************************************************************
    ' Name       :���������񂩂琮�����擾
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionToIntNo(ByVal strUseCr As String) As Integer
        Try
            Return CnvFractionToArray(strUseCr)(0)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :���������񂩂番�q���擾
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionToBunshi(ByVal strUseCr As String) As Integer
        Try
            Return CnvFractionToArray(strUseCr)(1)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :���������񂩂番����擾
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionToBunbo(ByVal strUseCr As String) As Integer
        Try
            Return CnvFractionToArray(strUseCr)(2)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :���������񂩂琮���E���q�E�����Integeer()���擾
    '
    '----------------------------------------------------------------------
    ' Comment    :�Ԃ����z��v�f��0:�����A1:���q�A2:����ł��B
    '             strUseCr��String.Empty�Ȃ炻�ꂼ��̒l��0�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/01/23
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFractionToArray(ByVal strUseCr As String) As Integer()
        Try
            strUseCr = strUseCr.Trim

            Dim intVal(2) As Integer

            If strUseCr <> String.Empty Then

                If strUseCr.IndexOf("."c) <> -1 Then
                    If CType(strUseCr, Double) = 0 Then
                        Return intVal '�S��0
                    Else
                        strUseCr = CnvFractionString(CType(strUseCr, Double))
                    End If
                End If

                If IsNumeric(strUseCr) = True Then
                    intVal(0) = CType(strUseCr, Integer)
                Else
                    '�����ݒ�
                    intVal(2) = CType(strUseCr.Split("/"c)(1), Integer)
                    '�����ƕ��q�̕����񂾂����擾
                    strUseCr = strUseCr.Split("/"c)(0)
                    If IsNumeric(strUseCr) = True Then
                        intVal(1) = CType(strUseCr, Integer)
                    Else
                        intVal(1) = CType(strUseCr.Split("��"c)(1), Integer)
                        intVal(0) = CType(strUseCr.Split("��"c)(0), Integer)
                    End If
                End If
            End If

            Return intVal

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�[������
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....���ɂȂ鐔�l 
    '�@�@�@�@�@�@ enmRoundPosition....�[���������鏬���_���ʒu
    '�@�@�@�@�@�@ enmRoundDiv.........enmRoundPosition�ʒu�̏����l�ɑ΂��鏈��
    '�@�@�@�@�@�@ ��)CnvRound(1.256,enmRoundPosition.P2,enmRoundDiv.Rounded) 
    '                �Ǝw�肵���ꍇ�͏�����2�ʂ��l�̌ܓ��Ȃ̂Ō��ʒl��1.3(1.3)�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/02/05
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvRound(ByVal dblValue As Double _
                                  , ByVal enmRoundPosition As ClsCst.enmRoundPosition _
                                  , ByVal enmRoundDiv As ClsCst.enmRoundDiv) As Double
        Try
            Select Case enmRoundDiv
                Case ClsCst.enmRoundDiv.Cutoff '�؂�̂�
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            Return Math.Floor(dblValue)
                        Case ClsCst.enmRoundPosition.P2
                            Return Math.Floor(dblValue * 10) / 10
                        Case ClsCst.enmRoundPosition.P3
                            Return Math.Floor(dblValue * 100) / 100
                        Case ClsCst.enmRoundPosition.P4
                            Return Math.Floor(dblValue * 1000) / 1000
                        Case ClsCst.enmRoundPosition.P5
                            Return Math.Floor(dblValue * 10000) / 10000
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            Return Math.Floor(dblValue * 100000000000000) / 100000000000000
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select

                Case ClsCst.enmRoundDiv.RoundUp '�؂�グ
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            Return Math.Round(dblValue + 0.4, 0, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P2
                            Return Math.Round(dblValue + 0.04, 1, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P3
                            Return Math.Round(dblValue + 0.004, 2, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P4
                            Return Math.Round(dblValue + 0.0004, 3, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P5
                            Return Math.Round(dblValue + 0.00004, 4, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            Return Math.Round(dblValue + 0.000000000000004, 4, MidpointRounding.AwayFromZero)
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select

                Case ClsCst.enmRoundDiv.Rounded '�l�̌ܓ�
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            Return Math.Round(dblValue, 0, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P2
                            Return Math.Round(dblValue, 1, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P3
                            Return Math.Round(dblValue, 2, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P4
                            Return Math.Round(dblValue, 3, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P5
                            Return Math.Round(dblValue, 4, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            Return Math.Round(dblValue, 14, MidpointRounding.AwayFromZero)
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�[������
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....���ɂȂ鐔�l 
    '�@�@�@�@�@�@ enmRoundPosition....�[���������鏬���_���ʒu
    '�@�@�@�@�@�@ enmRoundDiv.........enmRoundPosition�ʒu�̏����l�ɑ΂��鏈��
    '�@�@�@�@�@�@ ��)CnvRound(1.256,enmRoundPosition.P2,enmRoundDiv.Rounded) 
    '                �Ǝw�肵���ꍇ�͏�����2�ʂ��l�̌ܓ��Ȃ̂Ō��ʒl��1.3(1.3)�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/02/05
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvRound(ByVal dmlValue As Decimal _
                                  , ByVal enmRoundPosition As ClsCst.enmRoundPosition _
                                  , ByVal enmRoundDiv As ClsCst.enmRoundDiv) As Decimal
        Try
            Dim dmlNewValue As Decimal = dmlValue

            Select Case enmRoundDiv
                Case ClsCst.enmRoundDiv.Cutoff '�؂�̂�
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            dmlNewValue = Math.Floor(dmlValue)
                        Case ClsCst.enmRoundPosition.P2
                            dmlNewValue = Math.Floor(dmlValue * 10) / 10
                        Case ClsCst.enmRoundPosition.P3
                            dmlNewValue = Math.Floor(dmlValue * 100) / 100
                        Case ClsCst.enmRoundPosition.P4
                            dmlNewValue = Math.Floor(dmlValue * 1000) / 1000
                        Case ClsCst.enmRoundPosition.P5
                            dmlNewValue = Math.Floor(dmlValue * 10000) / 10000
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            dmlNewValue = Math.Floor(dmlValue * 100000000000000) / 100000000000000
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select

                Case ClsCst.enmRoundDiv.RoundUp '�؂�グ
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.4, 0, MidpointRounding.AwayFromZero), Decimal)
                        Case ClsCst.enmRoundPosition.P2
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.04, 1, MidpointRounding.AwayFromZero), Decimal)
                        Case ClsCst.enmRoundPosition.P3
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.004, 2, MidpointRounding.AwayFromZero), Decimal)
                        Case ClsCst.enmRoundPosition.P4
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.0004, 3, MidpointRounding.AwayFromZero), Decimal)
                        Case ClsCst.enmRoundPosition.P5
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.00004, 4, MidpointRounding.AwayFromZero), Decimal)
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.000000000000004, 4, MidpointRounding.AwayFromZero), Decimal)
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select

                Case ClsCst.enmRoundDiv.Rounded '�l�̌ܓ�
                    Select Case enmRoundPosition
                        Case ClsCst.enmRoundPosition.P1
                            dmlNewValue = Math.Round(dmlNewValue, 0, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P2
                            dmlNewValue = Math.Round(dmlNewValue, 1, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P3
                            dmlNewValue = Math.Round(dmlNewValue, 2, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P4
                            dmlNewValue = Math.Round(dmlNewValue, 3, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P5
                            dmlNewValue = Math.Round(dmlNewValue, 4, MidpointRounding.AwayFromZero)
                        Case ClsCst.enmRoundPosition.P15
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                            dmlNewValue = Math.Round(dmlNewValue, 14, MidpointRounding.AwayFromZero)
                            '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    End Select
            End Select

            Return dmlNewValue

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�؂�グ����
    '
    '----------------------------------------------------------------------
    ' Comment    :dmlValue.....���ɂȂ鐔�l 
    '�@�@�@�@�@�@ enmRoundPosition.........enmRoundPosition�ʒu�̏����l�ɑ΂��鏈��
    '�@�@�@�@�@�@ ��)CnvCutUp(1.250000000003,enmRoundPosition.P3) 
    '                �Ǝw�肵���ꍇ�͏�����2�ʂ�؂�グ�Ȃ̂Ō��ʒl��(1.26)�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :Y.Hayashi
    '        By  :2008/06/14
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvCutUp(ByVal dmlValue As Decimal _
                                  , ByVal enmRoundPosition As ClsCst.enmRoundPosition) As Decimal
        Try
            'Dim dblValue As Double = CType(dmlValue, Decimal)


            '�X�J���l�֐��@[f_get_RoundUP]�̌^ DECIMAL(13,5)���l������
            dmlValue = Math.Floor(dmlValue * 100000) / 100000

            Select Case enmRoundPosition
                Case ClsCst.enmRoundPosition.P1
                    '����
                    dmlValue = Math.Ceiling(dmlValue)
                Case ClsCst.enmRoundPosition.P2
                    '������P��
                    dmlValue = Math.Ceiling(dmlValue * 10) / 10

                Case ClsCst.enmRoundPosition.P3
                    '������Q��
                    dmlValue = Math.Ceiling(dmlValue * 100) / 100

                Case ClsCst.enmRoundPosition.P4
                    '������R��
                    dmlValue = Math.Ceiling(dmlValue * 1000) / 1000

                Case ClsCst.enmRoundPosition.P5
                    '������S��
                    dmlValue = Math.Ceiling(dmlValue * 10000) / 10000

                Case ClsCst.enmRoundPosition.P15
                    '������P�T��
                    '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    dmlValue = Math.Ceiling(dmlValue * 100000000000000) / 100000000000000
                    '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
            End Select

            Return dmlValue
            'Return CType(dblValue, Decimal)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�؂�グ����(�I�[�o�[���[�h) :��In Double�� Out Decimal
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....���ɂȂ鐔�l (Double)
    '�@�@�@�@�@�@ enmRoundPosition.........enmRoundPosition�ʒu�̏����l�ɑ΂��鏈��
    '�@�@�@�@�@�@ ��)CnvCutUp(1.250000000003,enmRoundPosition.P3) 
    '                �Ǝw�肵���ꍇ�͏�����2�ʂ�؂�グ�Ȃ̂Ō��ʒl��(1.26)�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :Y.Hayashi
    '        By  :2008/06/14
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvCutUp(ByVal dblValue As Double _
                                  , ByVal enmRoundPosition As ClsCst.enmRoundPosition) As Decimal
        Try
            '�X�J���l�֐��@[f_get_RoundUP]�̌^ DECIMAL(13,5)���l������
            dblValue = Math.Floor(dblValue * 100000) / 100000

            Select Case enmRoundPosition
                Case ClsCst.enmRoundPosition.P1
                    '����
                    dblValue = Math.Ceiling(dblValue)
                Case ClsCst.enmRoundPosition.P2
                    '������P��
                    dblValue = Math.Ceiling(dblValue * 10) / 10

                Case ClsCst.enmRoundPosition.P3
                    '������Q��
                    dblValue = Math.Ceiling(dblValue * 100) / 100

                Case ClsCst.enmRoundPosition.P4
                    '������R��
                    dblValue = Math.Ceiling(dblValue * 1000) / 1000

                Case ClsCst.enmRoundPosition.P5
                    '������S��
                    dblValue = Math.Ceiling(dblValue * 10000) / 10000

                Case ClsCst.enmRoundPosition.P15
                    '������P�T��
                    '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
                    dblValue = Math.Ceiling(dblValue * 100000000000000) / 100000000000000
                    '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
            End Select

            Return CType(dblValue, Decimal)

        Catch ex As Exception
            Throw
        End Try
    End Function


    '**********************************************************************
    ' Name       :�[������
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....���ɂȂ鐔�l 
    '�@�@�@�@�@�@ enmCutUpDiv.........enmCutUpDiv�ʒu�̏����l�ɑ΂��鏈��
    '�@�@�@�@�@�@ ��)CnvCutUp(1.0001,enmCutUpDiv.RoundUp) 
    '                �Ǝw�肵���ꍇ��2.0�ɂȂ�܂��B
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2008/02/05
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    'Public Shared Function CnvCutUp(ByVal dmlValue As Decimal _
    '                              , ByVal enmCutUpDiv As ClsCst.enmCutUpDiv) As Decimal
    '    Try
    '        Select Case enmCutUpDiv
    '            Case ClsCst.enmCutUpDiv.CutOff  '�������؎̂�
    '                Return Math.Floor(dmlValue)

    '            Case ClsCst.enmCutUpDiv.RoundUp '�������؂�グ
    '                Return Math.Ceiling(dmlValue)
    '        End Select

    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function


    '**********************************************************************
    ' Name       :�g�p�ʁA�p�����A���Z�l����A�P�ʐ��ʂ������_�\�L�̎����ɕϊ����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :��blnEasyScrRate�ɂ́ASysInf.EasyScrRate��ݒ肵�ĉ������B
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :FMS)S.Yamada
    '        By  :2008/02/15
    '----------------------------------------------------------------------
    ' Update On  :AST)M.Yoshida
    '        By  :2008/02/26
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvCrVal(ByVal dmlUseVal As Decimal, _
                                    ByVal dmlScrRate As Decimal, _
                                    ByVal dmlMPkGram As Decimal, _
                                    ByVal blnEasyScrRate As Boolean) As Decimal
        Dim dmlCrVal As Decimal = 0
        If dmlMPkGram = 0 Then Return 0
        If blnEasyScrRate Then
            If dmlScrRate <> 0 Then
                dmlCrVal = (dmlUseVal * (1 + (dmlScrRate / 100))) / dmlMPkGram
            Else
                If dmlMPkGram <> 0 Then
                    dmlCrVal = dmlUseVal / dmlMPkGram
                Else
                    Return 0
                End If
            End If
        Else
            If dmlScrRate <> 0 Then
                dmlCrVal = (dmlUseVal / (1 - (dmlScrRate / 100))) / dmlMPkGram
            Else
                If dmlMPkGram <> 0 Then
                    dmlCrVal = dmlUseVal / dmlMPkGram
                Else
                    Return 0
                End If
            End If
        End If

        Return Math.Round(dmlCrVal, 2, MidpointRounding.AwayFromZero)
    End Function

    '**********************************************************************
    ' Name       :�P�ʐ��ʁA�p�����A���Z�l����A�p�������������������g�p��(g)�ɕϊ����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :��blnEasyScrRate�ɂ́ASysInf.EasyScrRate��ݒ肵�ĉ������B
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :FMS)S.Yamada
    '        By  :2008/02/15
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvUseVal(ByVal StrCrVal As String, _
                                     ByVal dmlScrRate As Decimal, _
                                     ByVal dmlMPkGram As Decimal, _
                                     ByVal blnEasyScrRate As Boolean) As Decimal
        Dim dmlUseVal As Decimal = 0
        Dim dmlCrVal As Decimal = 0

        Dim intNo As Decimal = CnvFractionToArray(StrCrVal)(0)
        Dim intBunshi As Decimal = CnvFractionToArray(StrCrVal)(1)
        Dim intBunbo As Decimal = CnvFractionToArray(StrCrVal)(2)

        If intBunbo <> 0 Then
            dmlCrVal = intNo + intBunshi / intBunbo
        Else
            dmlCrVal = intNo
        End If

        If blnEasyScrRate Then
            dmlUseVal = (dmlCrVal * dmlMPkGram) / (1 + (dmlScrRate / 100))
        Else
            dmlUseVal = (dmlCrVal * dmlMPkGram) * (1 - (dmlScrRate / 100))
        End If

        Return Math.Round(dmlUseVal, 2, MidpointRounding.AwayFromZero)
    End Function

    '**********************************************************************
    ' Name       :�g�p�ʁA�p��������A�p���������܂񂾎g�p�ʂɕϊ����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :��blnEasyScrRate�ɂ́ASysInf.EasyScrRate��ݒ肵�ĉ������B
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yosihda
    '        By  :2008/02/17
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvScrVal(ByVal dblUseVal As Double, _
                                     ByVal dblScrRate As Double, _
                                     ByVal blnEasyScrRate As Boolean) As Double
        If blnEasyScrRate Then
            Return Math.Round(dblUseVal * (1 + (dblScrRate / 100)), 2, MidpointRounding.AwayFromZero)
        Else
            Return Math.Round(dblUseVal / (1 - (dblScrRate / 100)), 2, MidpointRounding.AwayFromZero)
        End If
    End Function

    '**********************************************************************
    ' Name       :�g�p�ʁA�p��������A�p���������܂񂾎g�p�ʂɕϊ����܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :��blnEasyScrRate�ɂ́ASysInf.EasyScrRate��ݒ肵�ĉ������B
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yosihda
    '        By  :2008/02/17
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvScrVal(ByVal dmlUseVal As Decimal, _
                                     ByVal dmlScrRate As Decimal, _
                                     ByVal blnEasyScrRate As Boolean) As Decimal
        If blnEasyScrRate Then
            Return Math.Round(dmlUseVal * (1 + (dmlScrRate / 100)), 2, MidpointRounding.AwayFromZero)
        Else
            Return Math.Round(dmlUseVal / (1 - (dmlScrRate / 100)), 2, MidpointRounding.AwayFromZero)
        End If
    End Function

    '**********************************************************************
    ' Name       :�����H�i�\���L�[��1�H�i�̕������Ԃ��܂��B
    '
    '----------------------------------------------------------------------
    ' Comment    :�H�i�R�[�h8���{�g�p��6+2���{�P�ʌv�Z�敪1���{����2���{���q2���{����2���{�^�󒲗�������2���{"_"
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yosihda
    '        By  :2008/06/02
    '----------------------------------------------------------------------
    ' Update On  :CM)Kuwahara
    '        By  :2011/02/17
    ' Comment    :�H�i�\���L�[�ɐ^�󒲗�������2����ǉ�
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function GetFdKey(ByVal strFdCd As String _
                                  , ByVal dmlUseVal As Decimal _
                                  , ByVal IntCrCalDiv As Integer _
                                  , ByVal IntNo As Integer _
                                  , ByVal IntBunShi As Integer _
                                  , ByVal IntBunBo As Integer _
                                  , ByVal IntVcCkDay As Integer) As String

        '�����H�i�\���L�[�i�H�i�R�[�h8���{�g�p��4+2���{�P�ʌv�Z�敪1���{����2���{���q2���{����2���{�^�󒲗�������2���{"_"�j
        Return strFdCd.PadRight(8).ToUpper _
             & dmlUseVal.ToString("0000.00").Remove(4, 1) _
             & IntCrCalDiv.ToString("0") _
             & IntNo.ToString("00") _
             & IntBunShi.ToString("00") _
             & IntBunBo.ToString("00") _
             & IntVcCkDay.ToString("00") _
             & "_"

    End Function

    '**********************************************************************
    ' Name       :�w��v�f���v�f�z��̒��ɑ��݂��Ă����True��Ԃ��B
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '�@�@�@�@�@�@ 
    '�@�@�@�@�@�@ 
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yosihda
    '        By  :2008/04/19
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function ChkIn(ByVal Value As Object, ByVal Values() As Object) As Boolean

        If Values Is Nothing OrElse Values.Length = 0 Then
            Return False
        Else
            Return Not (Array.IndexOf(Values, Value) = -1)
        End If

    End Function

    '**************************************************************************
    ' Name       :Const���擾
    '
    '--------------------------------------------------------------------------
    ' Comment    :
    '
    '--------------------------------------------------------------------------
    ' Design On  :2008/01/10
    '        By  :CM)Kuwahara
    '--------------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**************************************************************************
    Public Shared Function GetTblConst(ByVal aStr As String) As Col_GTblCst

        Dim objCol_GTblCst As Col_GTblCst
        Dim objRec_GTblCst As Rec_GTblCst

        Try

            ''�J���}��؂蕶�����z���Spit����
            Dim strSpit() As String = aStr.Split(CChar(","))
            ''�ʏ�̕����񂾂����ꍇ
            If strSpit.Length = 1 Then
                objCol_GTblCst = Nothing
            Else
                objCol_GTblCst = New Col_GTblCst

                '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
                For i As Integer = 0 To strSpit.Length - 1 Step 2
                    objRec_GTblCst = New Rec_GTblCst
                    objRec_GTblCst.ConstCode = CInt(strSpit(i))
                    objRec_GTblCst.ConstName = strSpit(i + 1)
                    objCol_GTblCst.Add(objRec_GTblCst)
                Next i
                '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^
            End If

            '----------------------------------------------
            '��O����
            '----------------------------------------------
        Catch ex As Exception
            objCol_GTblCst = Nothing
            Throw
        End Try

        Return objCol_GTblCst

    End Function

    '**********************************************************************
    ' Name       :�w�肵���t�H���_���ǂ߂�܂őҋ@����֐�
    '
    '----------------------------------------------------------------------
    ' Comment    :�w�肵���t�H���_���ǂ߂�܂őҋ@����֐�
    '             �q�`�r�ڑ���x���k�`�m�̏ꍇ�A�����Ȃ�b�n�o�x���ߓ�
    '�@�@�@�@�@�@ ���s����ƁA�G���[�ƂȂ��Ă��܂��ׁA���֐��őҋ@�B
    '
    '----------------------------------------------------------------------
    ' Parameter  :FileName: �Ď�����t�@�C��
    '             Interval: �Ď��Ԋu(ms)
    '             Times:    �Ď��J��Ԃ���
    '----------------------------------------------------------------------
    ' Return     :�t�@�C�����ǂ߂��ꍇTrue�A���Ԑ؂�Œ��߂��ꍇFalse
    '                              
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/04/23
    '        By  :FMS)N.Kasama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    : 
    '**********************************************************************
    Public Shared Function WaitForConnection(ByVal StrFileName As String, ByVal IntInterval As Integer, ByVal IntTimes As Integer) As Boolean

        Dim IntTM As Integer

        Try
            IntTM = 0

            Do Until IntTimes <= IntTM

                Try
                    If My.Computer.FileSystem.DirectoryExists(StrFileName) Then
                        WaitForConnection = True
                        Exit Function
                    End If
                Catch Ex As Exception
                    '�G���[�͖���
                End Try

                System.Threading.Thread.Sleep(CInt(IntInterval))

                IntTM = IntTM + 1

            Loop

            '���Ԑ؂ꎸ�s
            WaitForConnection = False

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :�w�肵���t�@�C�����ǂ߂�܂őҋ@����֐�
    '
    '----------------------------------------------------------------------
    ' Comment    :�w�肵���t�@�C�����ǂ߂�܂őҋ@����֐�
    '             �q�`�r�ڑ���x���k�`�m�̏ꍇ�A�����Ȃ�b�n�o�x���ߓ�
    '�@�@�@�@�@�@ ���s����ƁA�G���[�ƂȂ��Ă��܂��ׁA���֐��őҋ@�B
    '
    '----------------------------------------------------------------------
    ' Parameter  :FileName: �Ď�����t�@�C��
    '             Interval: �Ď��Ԋu(ms)
    '             Times:    �Ď��J��Ԃ���
    '----------------------------------------------------------------------
    ' Return     :�t�@�C�����ǂ߂��ꍇTrue�A���Ԑ؂�Œ��߂��ꍇFalse
    '                              
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/07/16
    '        By  :FMS)N.Kasama
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    : 
    '**********************************************************************

    Public Shared Function WaitFileRead(ByVal StrFileName As String, ByVal IntInterval As Integer, ByVal IntTimes As Integer) As Boolean

        Dim IntTM As Integer

        Try
            IntTM = 0

            Do Until IntTimes <= IntTM

                Try
                    If My.Computer.FileSystem.FileExists(StrFileName) Then
                        WaitFileRead = True
                        Exit Function
                    End If
                Catch Ex As Exception
                    '�G���[�͖���
                End Try

                System.Threading.Thread.Sleep(CInt(IntInterval))

                IntTM = IntTM + 1

            Loop

            '���Ԑ؂ꎸ�s
            WaitFileRead = False

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����Ԃ���A�P�O�̎��ԋ敪��Ԃ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  K.Taguchi
    '        By  :2008/07/03
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFormerEatTmDiv(ByVal intEatTmDiv As Integer) As Integer
        Try

            Select Case intEatTmDiv
                Case ClsCst.enmEatTimeDiv.MorningDiv : Return ClsCst.enmEatTimeDiv.NightDiv
                Case ClsCst.enmEatTimeDiv.AM10Div : Return ClsCst.enmEatTimeDiv.MorningDiv
                Case ClsCst.enmEatTimeDiv.LunchDiv : Return ClsCst.enmEatTimeDiv.AM10Div
                Case ClsCst.enmEatTimeDiv.PM15Div : Return ClsCst.enmEatTimeDiv.LunchDiv
                Case ClsCst.enmEatTimeDiv.EveningDiv : Return ClsCst.enmEatTimeDiv.PM15Div
                Case ClsCst.enmEatTimeDiv.NightDiv : Return ClsCst.enmEatTimeDiv.EveningDiv
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����Ԃ���A�P�O�̔N������Ԃ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  K.Taguchi
    '        By  :2008/07/03
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvFormerYMD(ByVal strYMD As String, ByVal intEatTmDiv As Integer) As String
        Try

            '���ԋ敪�����H��������
            If intEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv Then
                Return DateTime.Parse(Format(CInt(strYMD), "0000/00/00")).AddDays(-1).ToString("yyyyMMdd")
            Else
                Return strYMD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����Ԃ���A�P��̎��ԋ敪��Ԃ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  K.Taguchi
    '        By  :2008/07/03
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvLatterEatTmDiv(ByVal intEatTmDiv As Integer) As Integer
        Try

            Select Case intEatTmDiv
                Case ClsCst.enmEatTimeDiv.MorningDiv : Return ClsCst.enmEatTimeDiv.AM10Div
                Case ClsCst.enmEatTimeDiv.AM10Div : Return ClsCst.enmEatTimeDiv.LunchDiv
                Case ClsCst.enmEatTimeDiv.LunchDiv : Return ClsCst.enmEatTimeDiv.PM15Div
                Case ClsCst.enmEatTimeDiv.PM15Div : Return ClsCst.enmEatTimeDiv.EveningDiv
                Case ClsCst.enmEatTimeDiv.EveningDiv : Return ClsCst.enmEatTimeDiv.NightDiv
                Case ClsCst.enmEatTimeDiv.NightDiv : Return ClsCst.enmEatTimeDiv.MorningDiv
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :�H�����Ԃ���A�P��̔N������Ԃ�
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  K.Taguchi
    '        By  :2008/07/03
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function CnvLatterYMD(ByVal strYMD As String, ByVal intEatTmDiv As Integer) As String
        Try

            '���ԋ敪���[�ԐH��������
            If intEatTmDiv = ClsCst.enmEatTimeDiv.NightDiv Then
                Return DateTime.Parse(Format(CInt(strYMD), "0000/00/00")).AddDays(1).ToString("yyyyMMdd")
            Else
                Return strYMD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '�� Add by K.Kojima 2008/07/16 Start�i�P�D�I�����Ƀo�b�N�A�b�v�@�\�𓮍삳����@�\��ǉ��j*************************************************************************
    '**********************************************************************
    ' Name       :XML�t�@�C���w�荀�ڎ擾����
    '
    '----------------------------------------------------------------------
    ' Comment    :strFileName   -   XML�t�@�C���i�t���p�X�w��j
    '             strSection    -   �Z�N�V������
    '             strKey        -   �L�[��
    '
    '----------------------------------------------------------------------
    ' Design On  :2008/07/16
    '        By  :FMS)K.Kojima
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function GetXMLItem(ByVal strFileName As String, _
                                      ByVal strSection As String, _
                                      ByVal strKey As String) As String
        Try

            Dim objXmlDocument As New Xml.XmlDocument

            'XML�t�@�C����ǂݍ��ށB
            objXmlDocument.Load(strFileName)

            '�w��Z�N�V�����̃m�[�h���X�g���擾
            Dim objNodeList As Xml.XmlNodeList

            objNodeList = objXmlDocument.GetElementsByTagName(strSection)

            If Not objNodeList Is Nothing Then
                '�����Z�N�V���������������ꍇ�擪�̃m�[�h���g�p����
                Dim childNodeList As Xml.XmlNodeList
                childNodeList = objNodeList.Item(0).ChildNodes
                If Not childNodeList Is Nothing Then
                    '�Z�N�V�����̒��̃m�[�h�i�w��L�[�j������
                    Dim objNode As Xml.XmlNode
                    For Each objNode In childNodeList
                        If objNode.Name = strKey Then
                            '�ݒ�l�擾���߂�l�Ƃ��ĕԂ�
                            Return objNode.InnerText
                        End If
                    Next
                End If
            End If
            Return String.Empty
        Catch ex As Exception
            Return String.Empty
            Throw
        End Try

    End Function
    '�� Add by K.Kojima 2008/07/16 End  �i�P�D�I�����Ƀo�b�N�A�b�v�@�\�𓮍삳����@�\��ǉ��j*************************************************************************


    '**********************************************************************
    ' Name       :�t�@�C���o�C�i����r
    '
    '----------------------------------------------------------------------
    ' Comment    :�������e�Ȃ�True��Ԃ��܂��B
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/01/21
    '        By  :ATSP)M.Yosihda
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function FileCompare(ByVal strPath1 As String, ByVal strPath2 As String) As Boolean
        Try
            Using fs1 As System.IO.FileStream = System.IO.File.OpenRead(strPath1)
                Using fs2 As System.IO.FileStream = System.IO.File.OpenRead(strPath2)
                    If (fs1.Length <> fs2.Length) Then Return False
                    '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
                    While (fs1.Position < fs1.Length)
                        If fs1.ReadByte() <> fs2.ReadByte() Then Return False
                    End While
                    '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^
                End Using
            End Using

            Return True

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Summary    : �N��v�Z
    '----------------------------------------------------------------------
    ' Method     : Public Function GetPerAge
    '
    ' Parameter  : ByVal strBirth As String....���t������
    '�@�@�@�@�@�@�@ Optional ByVal enumDateInterval As DateInterval = DateInterval.Year
    ' Return     :  
    '
    '----------------------------------------------------------------------
    ' Comment    : �����œn���ꂽ�a�����Ƃ��Ă̓��t������(�A���W���̐��l����
    '              IsDate�֐���True�ƂȂ���t������)�ƃ��[�J���V�X�e�����t������
    '              �{�����_�̔N��l(��Q�����ɂ���Ēl�͕ς��܂��B)��Ԃ��܂��B
    '              �a�����t���s���l��Nothing�A�������̏ꍇ��-1��Ԃ��܂��B
    '              ��2/29���܂�ŃV�X�e���N���[�N�łȂ��l��3/01���܂�Ƃ��Čv�Z���܂��B�y���⑫�z�Q��
    '----------------------------------------------------------------------
    ' Design On  : 2008/03/03
    '        By  : FMS)H.Sakamoto
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function GetPerAge(ByVal strBirth As String _
                           , Optional ByVal enumDateInterval As DateInterval = DateInterval.Year _
                           , Optional ByVal strBaseDate As String = Nothing) As Long
        Try
            '---------------------------------------------------------------
            '�a�����̓��t����
            '---------------------------------------------------------------
            '�f�[�^�������Ă��Ȃ��ꍇ
            If strBirth Is Nothing OrElse strBirth.Trim() = "" Then Return -1
            '�W�����̏ꍇ�͐���W�����Ƃ݂Ȃ��āA���t������ɕϊ����܂��B
            If strBirth.Length = 8 AndAlso IsNumeric(strBirth) Then strBirth = CType(strBirth, Integer).ToString("0000/00/00")
            '���݂��Ȃ����t������̏ꍇ
            If IsDate(strBirth) = False Then Return -1
            '�L���ȓ��t������̏ꍇ
            Dim dteWork As Date = Date.Parse(strBirth)

            '---------------------------------------------------------------
            '����̓��t����
            '---------------------------------------------------------------
            '����t
            Dim dteBaseDate As Date
            '�f�[�^�������Ă��Ȃ��ꍇ
            If strBaseDate Is Nothing Then
                dteBaseDate = System.DateTime.Today
            Else
                '�W�����̏ꍇ�͐���W�����Ƃ݂Ȃ��āA���t������ɕϊ����܂��B
                If strBaseDate.Length = 8 AndAlso IsNumeric(strBaseDate) Then strBaseDate = CType(strBaseDate, Integer).ToString("0000/00/00")
                '���݂��Ȃ����t������̏ꍇ
                If IsDate(strBaseDate) = False Then Return -1
                '�L���ȓ��t������̏ꍇ
                dteBaseDate = Date.Parse(strBaseDate)
            End If

            '����t�ɑ΂��Ė����ȍ~�̒a�����t(������)���w�肳��Ă����ꍇ��-1��Ԃ��܂��B
            If dteWork > dteBaseDate Then Return -1

            If dteWork.ToString("MMdd") = "0229" _
                AndAlso dteBaseDate.Month = 2 _
                AndAlso DateTime.IsLeapYear(dteBaseDate.Year) = False Then
                '2/29���܂�̐l���A�����2�����A��N���[�N�łȂ��ꍇ��3/1���܂�Ƃ��܂��B�y���⑫�z�Q��
                dteWork = dteWork.AddDays(1)
            End If

            '�a���N�����Ɗ���t�Ƃœ��t�Ԋu�����߂܂��B
            Dim intRetValue As Long = DateDiff(enumDateInterval, dteWork, dteBaseDate)

            'DateInterval���N���ƌ����Ŏw�肳��Ă���ꍇ�ɂ͒a�������}�������ǂ����œ��t�Ԋu��␳���܂��B
            Select Case enumDateInterval
                Case DateInterval.Year
                    '��N�̌������A�a�������t�̌��������Ȃ�A�����ȍ~�ɒa�������}����ꍇ������̂ŔN���-1���܂��B
                    If dteBaseDate.ToString("MMdd") < dteWork.ToString("MMdd") Then intRetValue -= 1

                Case DateInterval.Month
                    '����̓����a�����t�̓������ŁA���A�����ȍ~�������������Ȃ�A����'�����ȍ~�̓�����'�ɒa�������}����ꍇ������̂Ō�����-1���܂��B
                    If dteBaseDate.ToString("dd") < dteWork.ToString("dd") _
                    AndAlso dteBaseDate.AddDays(1).Month = dteBaseDate.Month Then
                        intRetValue -= 1
                    End If
            End Select

            Return intRetValue

            '���y���⑫�z�N��v�Z�j�փX���@��������������������������������������������������������������������
            '���o�T: �t���[�S�Ȏ��T�w�E�B�L�y�f�B�A�iWikipedia�j�x�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '���N��v�Z�j�փX���@���i�˂�ꂢ��������ɂ��񂷂�ق���G����35�N12��2���@����50���j�́A�@�@��
            '�����{�̌��s�@���̈�ł���A�N��̌v�Z���@���߂�B���@�̕����@�̈�Ɉʒu�Â�����B�@�@�@��
            '���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '���{�@�ɂ��΁A�N��͏o���̓���肱����N�Z���i1���j�A��ɏ]���ĔN�����͌����ɂ���Čv�Z���� �@��
            '���i2���A���@143��1���B�Ȃ��A�N��̂ƂȂ����Ɋւ���@��1���j�B�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '�����Ȃ킿�A���N�A�a�����i�u�N�Z���j�����X�����v�j�̑O���������Ă���N��͖������邩��@�@�@�@�@��
            '���i�{�@2���A���@143��2���{���j�A�a�����̑O�����I������u�ԁi�a�����̌ߑO0��00���̒��O�j��1�΁@��
            '���������邱�ƂɂȂ�B�������A�a�����O�̍Ō�Ɏn�܂錎�ɒa���������݂��Ȃ��Ƃ��@�@�@�@�@�@�@�@�@��
            '���i�u�N���ȃe���ԃ��胁�^���ꍇ�j���e�Ō�m���j�������i�L�g�L�v�j�́A���̌��̖����������Ă��́@��
            '���N���������i�{�@2���A���@143��2���A���j�B�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '���@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '���Ⴆ�΁A2004�N10��4���ɏo�������҂̔N��́A�a�����̑O���ł��閈�N10��3�����I������u�ԁA�@�@ ��
            '�����Ȃ킿10��4���ߑO0��00���̒��O��1�΂�������B�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '���܂��A2004�N�i�[�N�j2��29���ɏo�������҂̔N��́A���N�ɂ͒a�����O�̍Ō�Ɏn�܂錎�ł���@�@�@��
            '��2���ɒa�����ł���2��29�������݂��Ȃ�����A���̌��̖����ł���2��28�����I������u�ԁA�@�@�@�@�@��
            '�����Ȃ킿3��1���ߑO0��00���̒��O�ɁA�[�N�ɂ͒a�����̑O���ł���2��28�����I������u�ԁA�@�@�@�@ ��
            '�����Ȃ킿2��29���ߑO0��00���̒��O�ɁA���ꂼ��1�΂�������킯�ł���B�@�@�@�@�@�@�@�@�@�@�@�@�@��
            '��������������������������������������������������������������������������������������������������

            '--------------------
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :������̎w��o�C�g��������
    '
    '----------------------------------------------------------------------
    ' Comment    :�S�p���܂ޕ�������w��o�C�g�ȉ��ɕ������܂��B
    '           �@�Ⴆ�΁ASplitString("123������",4)��{"123","����","��"}��Ԃ��܂��B
    '----------------------------------------------------------------------
    ' Design On  :2012/05/28
    '        By  :ATSP)M.Yoshida
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function SplitString(ByVal strText As String, ByVal intSplitLength As Integer) As String()
        Try
            If strText.Trim = String.Empty Then Return New String() {}

            '�������ꕶ���Â�������20�o�C�g�ȉ���Rec��������Col�ɒǉ����܂��B
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
            Dim intI As Integer = 0
            Dim strExamRst As String = ""
            Dim Col As New Collections.Specialized.StringCollection
            '�^�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�P�_
            Do
                'intI��������̍Ōォ�AstrExamRst��intI�ʒu��1������A�����Ďw��o�C�g�ȏ�ɂȂ�Ȃ�Rec��ǉ����܂�
                If intI > (strText.Length - 1) _
                OrElse (enc.GetBytes(strExamRst & strText.Substring(intI, 1)).Length > intSplitLength) Then

                    'Col�ɒǉ�
                    Col.Add(strExamRst)

                    '����Rec�ׂ̈ɃN���A
                    strExamRst = ""
                End If

                'intI���������𒴂�����J��Ԃ��𔲂��܂��B
                If intI > (strText.Length - 1) Then Exit Do

                'strExamRst�ɕ����ʒu�̎���1������A�����܂��B
                strExamRst &= strText.Substring(intI, 1)

                '���̕����ʒu�ɃJ�E���g�A�b�v
                intI += 1
            Loop
            '�_�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�Q�^

            '�����z��ɕԂ��܂��B
            Dim strRetR(Col.Count - 1) As String
            Col.CopyTo(strRetR, 0)

            '�Ԃ��܂��B
            Return strRetR

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :AddPathSeparatorChar
    '
    '----------------------------------------------------------------------
    ' Comment    :��؂蕶����t������
    '
    '----------------------------------------------------------------------
    ' Design On  :2014/03/11
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function AddPathSeparatorChar(ByVal strPath As String, _
                                       Optional ByVal strSeparator As String = "\") As String

        Dim strRtnPath As String

        Try

            strRtnPath = IIf(strPath.EndsWith(strSeparator) = True, strPath, strPath & strSeparator).ToString()

            Return strRtnPath

        Catch ex As Exception

            Return strPath

        End Try

    End Function

    '**********************************************************************
    ' Name       :SplitDataByChar
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTX)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function SplitDataByChar(ByVal strVal As String, _
                                           ByVal strSplitChar As String) As String()

        Dim objRtnCol As New Collections.Specialized.StringCollection

        Try

            ''��؂蕶�����Ȃ��ꍇ
            If strVal.IndexOf(strSplitChar) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplitChar, Char))

                For Each strWkVal As String In strRtns
                    objRtnCol.Add(strWkVal)
                Next

            Else
                If strVal.Trim <> "" Then
                    objRtnCol.Add(strVal)
                End If
            End If

            Dim strRetR(objRtnCol.Count - 1) As String
            objRtnCol.CopyTo(strRetR, 0)

            '�Ԃ��܂��B
            Return strRetR

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetSplitDataByIndex
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTX)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Shared Function GetSplitDataByIndex(ByVal strVal As String, _
                                               ByVal strSplit As String, _
                                               ByVal intIndex As Integer) As String

        Dim strRtn As String = ""

        Try

            Dim objStrCol As New Collections.Specialized.StringCollection

            ''��؂蕶�����Ȃ��ꍇ
            If strVal.IndexOf(strSplit) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit, Char))

                For Each strWkVal As String In strRtns
                    objStrCol.Add(strWkVal)
                Next

            Else
                If strVal.Trim <> "" Then
                    objStrCol.Add(strVal)
                End If
            End If

            If intIndex < objStrCol.Count Then

                strRtn = objStrCol.Item(intIndex)

            End If

            Return strRtn

        Catch ex As Exception

            Return ""

        End Try

    End Function

End Class
