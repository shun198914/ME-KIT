'Option Strict On
''**************************************************************************
'' Name       :�y�h�o�t�@�C������N���Xtestx
''
''--------------------------------------------------------------------------
'' Comment    :
''
''--------------------------------------------------------------------------
'' Design On  :2008/03/05
''        By  :FMS)Y.Maki
''--------------------------------------------------------------------------
'' Update On  :
''        By  :
'' Comment    :
''**************************************************************************
''**************************************************************************
'' Class
''**************************************************************************
'Public Class ClsZip
'    '**************************************************************************
'    ' Inner Class
'    '**************************************************************************

'    '**********************************************************************
'    ' �o�u���b�N�C�x���g
'    '**********************************************************************
'    Public Event OperationInProgress(ByVal isTopItem As Boolean, ByVal TagetPath As String, ByVal Now As Integer, ByVal total As Integer, ByRef Cancel As Boolean)

'    Public Class ClsZipFile
'        '**********************************************************************
'        ' �v���C�x�[�g�ϐ�
'        '**********************************************************************
'        Private m_strFileName As String ' �t�@�C����
'        Private m_lngFileSize As Long ' �t�@�C���T�C�Y
'        Private m_lngCompressedSize As Long ' ���k�t�@�C���T�C�Y
'        Private m_dteTimeStamp As Date ' �^�C���X�^���v

'        '**********************************************************************
'        ' �v���C�x�[�g�ϐ�
'        '**********************************************************************

'        '**********************************************************************
'        ' Name       :�t�@�C�����v���p�e�B
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Property FileName() As String
'            Get
'                Return m_strFileName
'            End Get
'            Set(ByVal value As String)
'                m_strFileName = value
'            End Set
'        End Property

'        '**********************************************************************
'        ' Name       :�t�@�C���T�C�Y�v���p�e�B
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Property FileSize() As Long
'            Get
'                Return m_lngFileSize
'            End Get
'            Set(ByVal value As Long)
'                m_lngFileSize = value
'            End Set
'        End Property

'        '**********************************************************************
'        ' Name       :���k�t�@�C���T�C�Y�v���p�e�B
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Property CompressedSize() As Long
'            Get
'                Return m_lngCompressedSize
'            End Get
'            Set(ByVal value As Long)
'                m_lngCompressedSize = value
'            End Set
'        End Property

'        '**********************************************************************
'        ' Name       :�t�@�C���^�C���X�^���v�v���p�e�B
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Property FileTimestamp() As DateTime
'            Get
'                Return m_dteTimeStamp
'            End Get
'            Set(ByVal value As DateTime)
'                m_dteTimeStamp = value
'            End Set
'        End Property
'    End Class

'    '**************************************************************************
'    ' Inner Class
'    '**************************************************************************
'    Public Class ClsZipFiles
'        '**********************************************************************
'        ' �v���C�x�[�g�ϐ�
'        '**********************************************************************
'        Private m_objZipFiles As Hashtable

'        '**********************************************************************
'        ' Name       :�R���X�g���N�^
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Sub New()
'            m_objZipFiles = New Hashtable()
'        End Sub

'        '**********************************************************************
'        ' Name       :�t�@�C����
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Function Count() As Integer
'            Try
'                Return m_objZipFiles.Count
'            Catch ex As Exception
'                Throw
'            End Try
'        End Function

'        '**********************************************************************
'        ' Name       :���k�t�@�C���R���N�V�����̓o�^
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Function Add(ByVal objZipFile As ClsZipFile) As Integer
'            Try
'                Dim intKey As Integer = m_objZipFiles.Count
'                m_objZipFiles.Add(intKey, objZipFile)
'                Return intKey
'            Catch ex As Exception
'                Throw
'            End Try
'        End Function

'        '**********************************************************************
'        ' Name       :���k�t�@�C���R���N�V�����̂P�G���g��
'        '
'        '----------------------------------------------------------------------
'        ' Comment    :
'        '
'        '----------------------------------------------------------------------
'        ' Design On  :2008/03/05
'        '        By  :FMS)Y.Maki
'        '----------------------------------------------------------------------
'        ' Update On  :
'        '        By  :
'        ' Comment    :
'        '**********************************************************************
'        Public Function Item(ByVal intKey As Integer) As ClsZipFile
'            Try
'                If (intKey >= 0) AndAlso (intKey < m_objZipFiles.Count) Then
'                    Return CType(m_objZipFiles.Item(intKey), ClsZipFile)
'                End If
'                Return Nothing
'            Catch ex As Exception
'                Throw
'            End Try
'        End Function
'    End Class

'    '**********************************************************************
'    ' Name       :���k�t�@�C���ꗗ�����X�g�o�͂���
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :2010/01/26���݁A�N���g�p���Ă��Ȃ��i�o�O�L��j�ׁ̈A
'    '             �g�p�֎~
'    '----------------------------------------------------------------------
'    ' Design On  :2008/03/05
'    '        By  :FMS)Y.Maki
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Public Function ListZipFiles(ByVal strZipPath As String) As ClsZipFiles
'        Try
'            Dim objZipFiles As New ClsZipFiles

'            '�J��ZIP�t�@�C���̐ݒ�
'            Dim zipPath As String = strZipPath

'            '�ǂݍ���
'            Dim objFileStream As New java.io.FileInputStream(zipPath)
'            Dim objZipInputStream As New java.util.zip.ZipInputStream(objFileStream)

'            'ZIP���̃t�@�C�������擾
'            Do While True
'                Dim objZipEntry As java.util.zip.ZipEntry = objZipInputStream.getNextEntry()
'                If (objZipEntry Is Nothing) Then
'                    Exit Do
'                End If
'                If Not objZipEntry.isDirectory() Then
'                    '����\������
'                    Dim objZipFile As New ClsZipFile()
'                    With objZipFile
'                        .FileName = objZipEntry.getName()
'                        .FileSize = objZipEntry.getSize()
'                        .CompressedSize = objZipEntry.getCompressedSize()
'                        .FileTimestamp = CType(New java.util.Date(objZipEntry.getTime()).toLocaleString(), DateTime)
'                    End With
'                    objZipFiles.Add(objZipFile)
'                End If
'                objZipInputStream.closeEntry()
'            Loop

'            '����
'            objZipInputStream.close()
'            objFileStream.close()

'            Return objZipFiles

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :�w�肳�ꂽ�t�@�C�������k����
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/03/05
'    '        By  :FMS)Y.Maki
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Public Function CompressZipFiles(ByVal strZipFileName As String _
'                                   , ByVal objZipFiles As ArrayList _
'                          , Optional ByVal blnIsFolderSave As Boolean = False) As Integer
'        Try
'            ' ��{�̈��k�����Ăяo���p�N���X�Ƀt�@�C�������Z�b�g����
'            Dim intCompressedFileCount As Integer = 0
'            Dim objTargetZipFiles As ClsZipFiles
'            objTargetZipFiles = New ClsZipFiles()
'            For intFileIndex As Integer = 0 To objZipFiles.Count - 1
'                Dim objTargetZipFile As ClsZipFile = New ClsZipFile()
'                objTargetZipFile.FileName = CType(objZipFiles(intFileIndex), String)
'                objTargetZipFiles.Add(objTargetZipFile)
'            Next

'            ' ��{�̈��k�������Ăяo��
'            Return CompressZipFiles(strZipFileName, objTargetZipFiles, blnIsFolderSave)
'        Catch ex As Exception
'            Throw
'        End Try
'    End Function
'    Public Function CompressZipFiles(ByVal strZipFileName As String _
'                                   , ByVal objZipFiles As ClsZipFiles _
'                          , Optional ByVal blnIsFolderSave As Boolean = False) As Integer
'        Dim intCompressedFileCount As Integer = 0

'        Dim objFileStream As java.io.FileOutputStream = Nothing
'        Dim objZipFileStream As java.util.zip.ZipOutputStream = Nothing
'        Dim blnCancel As Boolean = False

'        Try
'            ' �y�h�o�o�̓X�g���[���̍쐬
'            objFileStream = New java.io.FileOutputStream(strZipFileName)
'            objZipFileStream = New java.util.zip.ZipOutputStream(objFileStream)

'            Dim timer As New DateTime

'            ' ���ׂẴt�@�C���ɑ΂��ď�������
'            For intRecordIndex As Integer = 0 To objZipFiles.Count - 1

'                If timer.Second <> Now.Second _
'                OrElse intCompressedFileCount = 0 _
'                OrElse intCompressedFileCount = (objZipFiles.Count - 1) Then
'                    '�C�x���g���N�����܂��B
'                    RaiseEvent OperationInProgress(True, objZipFiles.Item(intRecordIndex).FileName, intCompressedFileCount + 1, objZipFiles.Count, blnCancel)

'                    '�L�����Z�����ꂽ�ꍇ�͌J��Ԃ��𔲂��܂��B
'                    If blnCancel = True Then Return -1

'                    timer = Now
'                End If

'                If System.IO.File.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' �t�@�C���̏ꍇ
'                    If CompressFile(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                ElseIf System.IO.Directory.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' �f�B���N�g���̏ꍇ
'                    If CompressDirectory(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                End If

'                ' ���k�t�@�C�����̍X�V
'                intCompressedFileCount += 1
'            Next

'            Return intCompressedFileCount


'        Catch ex As Exception
'            Throw
'        Finally
'            ' �y�h�o�o�̓X�g���[�������
'            If objZipFileStream IsNot Nothing Then
'                objZipFileStream.close()
'            End If
'            If objFileStream IsNot Nothing Then
'                objFileStream.close()
'            End If
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :???
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :???
'    '        By  :???
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :�Ȃ񂩂��̃��\�b�h�̓_�����Ǝv���܂��B
'    '**********************************************************************
'    Private Function CompressDirectory(ByVal objZipDirectory As ClsZipFile _
'                                     , ByVal objZipFileStream As java.util.zip.ZipOutputStream _
'                                     , ByVal blnIsFolderSave As Boolean) As Boolean
'        Try
'            Dim strFiles As String() = System.IO.Directory.GetFiles(objZipDirectory.FileName)

'            For Each strFilePath As String In strFiles
'                ' �t�@�C�������k����
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = strFilePath
'                CompressFile(objZipFile, objZipFileStream, blnIsFolderSave)
'            Next

'            If strFiles.Length = 0 Then
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = objZipDirectory.FileName
'                CompressFile(objZipFile, objZipFileStream, blnIsFolderSave)
'            End If

'            '�T�u�t�H���_����������i�w�肵���t�H���_�j
'            For Each strDirPath As String In System.IO.Directory.GetDirectories(objZipDirectory.FileName)
'                Dim objZipSubDirectory As ClsZipFile = New ClsZipFile()
'                objZipSubDirectory.FileName = strDirPath & "\"
'                CompressDirectoryRelativePath(objZipSubDirectory, objZipFileStream, blnIsFolderSave)
'            Next

'            '����
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :�t�@�C�����k
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/03/05
'    '        By  :FMS)Y.Maki
'    '----------------------------------------------------------------------
'    ' Update On  :2008/05/22
'    '        By  :FMS)N.Kasama
'    ' Comment    :�𓀌�̓��t���ُ�ɌÂ��Ȃ錻�ۂ����蒲�������Ƃ���
'    '             ���L�ŃG���g���̏C�����Ԃ�ݒ肵�Ă��邪�A�s�v�Ǝv��ꂽ�̂�
'    '             �R�����g�����琳��ȓ��t�ƂȂ���
'    ' Update On  :2008/05/28
'    '        By  :FMS)N.Kasama
'    ' Comment    :�u���[�J���w�b�_�[�̃f�[�^���s���ł��v�Ƃ������b�Z�[�W�Ή�
'    '**********************************************************************
'    Private Function CompressFile(ByVal objZipFile As ClsZipFile _
'                                , ByVal objZipFileStream As java.util.zip.ZipOutputStream _
'                                , ByVal blnIsFolderSave As Boolean) As Boolean
'        Try
'            ' �t�@�C�����̓X�g���[���̍쐬
'            Dim strFileName As String = objZipFile.FileName
'            If blnIsFolderSave = True Then ' �t�H���_���ێ�����ꍇ
'                ' ���[�g���폜���A\��/�ɒu�����Ă���
'                strFileName = strFileName.Remove(0, System.IO.Path.GetPathRoot(strFileName).Length).Replace("\", "/")
'            Else ' �t�@�C�����݂̂̏ꍇ
'                ' �t�@�C�����{�݂̂̂𒊏o����
'                strFileName = System.IO.Path.GetFileName(strFileName)
'            End If

'            ' �W�j�o�C�g�����̓t�@�C������ǂݍ��݁A�o�̓t�@�C���֏�����
'            Dim objZipEntry As New java.util.zip.ZipEntry(strFileName)
'            objZipEntry.setMethod(java.util.zip.ZipEntry.DEFLATED)

'            '��hayashi 2010/01/26
'            '���k���t�@�C���̍X�V���t���G�|�b�NTime�Őݒ肷��
'            '(�^�C���X�^���v���ݒ肳��Ă��Ȃ��ꍇ�͐ݒ肵�Ȃ�)
'            If objZipFile.FileTimestamp.ToOADate <> 0.0 Then
'                objZipEntry.setTime(toEpochTime(objZipFile.FileTimestamp))
'                '��hayashi 2010/01/26
'            End If
'            objZipFileStream.putNextEntry(objZipEntry)
'            Dim bytBuffer(8191) As System.SByte
'            Dim lngLength As Long = 0
'            Dim intCompressedLength As Integer = 0
'            Dim intReadedLength As Integer
'            Dim objFileInputStream As New java.io.FileInputStream(objZipFile.FileName)

'            '0529
'            Dim crc As New java.util.zip.CRC32

'            Do While True
'                intReadedLength = objFileInputStream.read(bytBuffer, 0, bytBuffer.Length)
'                If intReadedLength <= 0 Then
'                    Exit Do
'                End If

'                '0529
'                crc.update(bytBuffer, 0, intReadedLength)

'                objZipFileStream.write(bytBuffer, 0, intReadedLength)
'                lngLength += intReadedLength
'            Loop

'            '2008/05/22Dim lngTime As Long = java.util.Date.parse(String.Format("{0:yyyy/MM/dd HH:mm:ss}", System.IO.File.GetLastWriteTime(objZipFile.FileName)))
'            '2008/05/22objZipEntry.setTime(lngTime)
'            '2008/05/28objZipEntry.setSize(lngLength)
'            objZipEntry.setSize(lngLength)
'            objZipEntry.setCrc(crc.getValue())

'            ' ���̓X�g���[�������
'            objFileInputStream.close()

'            ' �y�h�o�G���g����ǉ�����
'            objZipFileStream.closeEntry()

'            '0529
'            objZipFileStream.flush()

'            '����
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :�w�肳�ꂽ���k�t�@�C�������ׂēW�J����
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/03/05
'    '        By  :FMS)Y.Maki
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Public Function ExpandZipAll(ByVal strSrcZipFileName As String, ByVal strDestExpandFolderName As String) As Integer
'        Try
'            '�ǂݍ���
'            Dim objFileStream As New java.io.FileInputStream(strSrcZipFileName)
'            Dim objZipStream As New java.util.zip.ZipInputStream(objFileStream)
'            If System.IO.Directory.Exists(strDestExpandFolderName) = False Then
'                Try
'                    System.IO.Directory.CreateDirectory(strDestExpandFolderName)
'                Catch
'                    Throw New Exception("�ȉ��̃t�H���_���쐬�ł��܂���ł����B" & ControlChars.CrLf & strDestExpandFolderName)
'                End Try
'            End If

'            Dim timer As New DateTime
'            Dim blnCancel As Boolean = False
'            Dim intI As Integer = 0
'            Dim dtFileTimestamp As Date
'            Dim strFileName As String = ""

'            'ZIP���̃t�@�C�������擾
'            Dim objZipEntry As java.util.zip.ZipEntry = objZipStream.getNextEntry()

'            Do Until objZipEntry Is Nothing

'                If Not objZipEntry.isDirectory() Then

'                    '�t�@�C����
'                    strFileName = objZipEntry.getName()

'                    If timer.Second <> Now.Second _
'                    OrElse intI = 0 Then
'                        '�C�x���g���N�����܂��B
'                        RaiseEvent OperationInProgress(True, strFileName, intI + 1, -1, blnCancel)

'                        '�L�����Z�����ꂽ�ꍇ�͌J��Ԃ��𔲂��܂��B
'                        If blnCancel = True Then Return -1

'                        timer = Now
'                    End If

'                    intI += 1

'                    '�W�J��̃p�X
'                    Dim strOutputFolderName As String = System.IO.Path.Combine(strDestExpandFolderName, System.IO.Path.GetDirectoryName(strFileName))
'                    If System.IO.Directory.Exists(strOutputFolderName) = False Then ' �t�H���_�����݂��Ȃ��ꍇ
'                        ' �t�H���_���쐬����
'                        System.IO.Directory.CreateDirectory(strOutputFolderName)
'                    End If
'                    Dim strOutputFileName As String = System.IO.Path.Combine(strOutputFolderName, System.IO.Path.GetFileName(strFileName))
'                    'FileOutputStream�̍쐬
'                    Dim objFileOutputStream As java.io.FileOutputStream
'                    Try
'                        objFileOutputStream = New java.io.FileOutputStream(strOutputFileName)
'                    Catch
'                        Throw New Exception("�ȉ��̃t�@�C�����쐬�ł��܂���ł����B" & ControlChars.CrLf & strOutputFileName _
'                                           & (intI - 1).ToString & "�̃t�@�C���͓W�J����Ă��܂��B")
'                    End Try
'                    '������
'                    Dim bytBuffer(8191) As System.SByte

'                    Do While True
'                        Dim len As Integer = objZipStream.read(bytBuffer, 0, bytBuffer.Length)
'                        If len <= 0 Then
'                            Exit Do
'                        End If
'                        objFileOutputStream.write(bytBuffer, 0, len)
'                    Loop

'                    ' �o�̓X�g���[�������
'                    objFileOutputStream.close()
'                    objFileOutputStream.Finalize()
'                    objFileOutputStream = Nothing

'                    bytBuffer = New System.SByte() {}
'                    bytBuffer = Nothing

'                    '�G�|�b�NTime��ϊ����čX�V�����i�[
'                    dtFileTimestamp = fromEpochTime(objZipEntry.getTime())

'                    ' �^�C���X�^���v�����k�t�@�C���̏��̒u������
'                    System.IO.File.SetCreationTime(strOutputFileName, dtFileTimestamp)

'                    System.IO.File.SetLastWriteTime(strOutputFileName, dtFileTimestamp)

'                    System.IO.File.SetLastAccessTime(strOutputFileName, dtFileTimestamp)
'                End If

'                objZipStream.closeEntry()

'                'ZIP���̃t�@�C�������擾
'                objZipEntry = objZipStream.getNextEntry()
'            Loop

'            If intI > 0 Then
'                '�C�x���g���N�����܂��B
'                RaiseEvent OperationInProgress(True, strFileName, intI + 1, -1, blnCancel)
'            End If

'            ' ���̓X�g���[�������()
'            objZipStream.close()
'            objFileStream.close()

'            Return intI

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '�� Add by K.Kojima 2008/08/26 Start�i�P�DFrmRestoreTool.exe�����s����t�H���_���AGTX�t�H���_����GTX_Data�ɍ쐬�����e���|�����t�H���_�Ŏ��s����l�ɕύX�j*************************************************************************
'#Region "�w��f�B���N�g�����k�����i���΃p�X�j�@CompressZipFilesRelativePath"

'    '**********************************************************************
'    ' Name       :�w��f�B���N�g�����k�����i���΃p�X�j
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :�w�肳�ꂽ�f�B���N�g���𑊑΃p�X�ň��k�������s��
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :strZipFileName    -   ���k����ZIP�t�@�C���p�X���i�t���p�X�j
'    '             objZipFiles       -   ���k����f�B���N�g��
'    '             blnIsFolderSave   -   True    =   �t�H���_���ێ�����ꍇ
'    '                                   False   =   �t�H���_���ێ����Ȃ��ꍇ
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Integer           -   ���k�t�@�C����
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/08/26
'    '        By  :FMS)K.Kojima
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Public Function CompressZipFilesRelativePath(ByVal strZipFileName As String, _
'                                                 ByVal objZipFiles As ClsZipFiles, _
'                                        Optional ByVal blnIsFolderSave As Boolean = False) As Integer
'        Try
'            Dim intCompressedFileCount As Integer = 0

'            Dim objFileStream As java.io.FileOutputStream = Nothing
'            Dim objZipFileStream As java.util.zip.ZipOutputStream = Nothing
'            Dim blnCancel As Boolean = False

'            Try
'                ' �y�h�o�o�̓X�g���[���̍쐬
'                objFileStream = New java.io.FileOutputStream(strZipFileName)
'                objZipFileStream = New java.util.zip.ZipOutputStream(objFileStream)

'                Dim timer As New DateTime

'                ' ���ׂẴt�@�C���ɑ΂��ď�������
'                For intRecordIndex As Integer = 0 To objZipFiles.Count - 1

'                    If timer.Second <> Now.Second _
'                    OrElse intCompressedFileCount = 0 _
'                    OrElse intCompressedFileCount = (objZipFiles.Count - 1) Then
'                        '�C�x���g���N�����܂��B
'                        RaiseEvent OperationInProgress(True, objZipFiles.Item(intRecordIndex).FileName, intCompressedFileCount + 1, objZipFiles.Count, blnCancel)

'                        '�L�����Z�����ꂽ�ꍇ�͌J��Ԃ��𔲂��܂��B
'                        If blnCancel = True Then Return -1

'                        timer = Now
'                    End If

'                    If System.IO.File.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' �t�@�C���̏ꍇ
'                        If CompressFileRelativePath(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                    ElseIf System.IO.Directory.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' �f�B���N�g���̏ꍇ
'                        If CompressDirectoryRelativePath(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave, objZipFiles.Item(intRecordIndex).FileName) = False Then Return -1
'                    End If

'                    ' ���k�t�@�C�����̍X�V
'                    intCompressedFileCount += 1
'                Next

'                '����
'                Return intCompressedFileCount

'            Catch ex As Exception
'                Throw
'            Finally
'                ' �y�h�o�o�̓X�g���[�������
'                If objZipFileStream IsNot Nothing Then
'                    objZipFileStream.close()
'                End If
'                If objFileStream IsNot Nothing Then
'                    objFileStream.close()
'                End If
'            End Try

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'#End Region

'#Region "�f�B���N�g�����k�����i���΃p�X�j�@CompressDirectoryRelativePath"

'    '**********************************************************************
'    ' Name       :�f�B���N�g�����k�����i���΃p�X�j
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :�w�肳�ꂽ�f�B���N�g���ɑ��݂���t�@�C�������k���A
'    '             �t�H���_�����݂����ꍇ�͍ċA���A���k�������s��
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :objZipDirectory   -   ���k����f�B���N�g��
'    '             objZipFileStream  -   ���k�`��
'    '             blnIsFolderSave   -   True    =   �t�H���_���ێ�����ꍇ
'    '                                   False   =   �t�H���_���ێ����Ȃ��ꍇ
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Boolean           -   ���g�p
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/08/26
'    '        By  :FMS)K.Kojima
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Private Function CompressDirectoryRelativePath(ByVal objZipDirectory As ClsZipFile, _
'                                                   ByVal objZipFileStream As java.util.zip.ZipOutputStream, _
'                                                   ByVal blnIsFolderSave As Boolean, _
'                                          Optional ByVal strTopDirectory As String = "") As Boolean
'        Try
'            Dim strFiles As String() = System.IO.Directory.GetFiles(objZipDirectory.FileName)

'            Dim timer As New DateTime
'            Dim strFilePath As String
'            Dim blnCancel As Boolean = False

'            For intI As Integer = 0 To strFiles.Length - 1

'                strFilePath = strFiles(intI)

'                If timer.Second <> Now.Second _
'                OrElse intI = 0 _
'                OrElse intI = (strFiles.Length - 1) Then
'                    '�C�x���g���N�����܂��B
'                    RaiseEvent OperationInProgress(False, objZipDirectory.FileName, intI + 1, strFiles.Length, blnCancel)

'                    '�L�����Z�����ꂽ�ꍇ�͌J��Ԃ��𔲂��܂��B
'                    If blnCancel = True Then Return False

'                    timer = Now
'                End If

'                ' �t�@�C�������k����
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = strFilePath
'                objZipFile.FileTimestamp = System.IO.File.GetLastWriteTime(strFilePath)
'                If CompressFileRelativePath(objZipFile, objZipFileStream, blnIsFolderSave, strTopDirectory) = False Then Return False
'            Next

'            timer = Nothing

'            If strFiles.Length = 0 Then
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = objZipDirectory.FileName
'                objZipFile.FileTimestamp = System.IO.File.GetLastWriteTime(objZipDirectory.FileName)
'                If CompressFileRelativePath(objZipFile, objZipFileStream, blnIsFolderSave, strTopDirectory) = False Then Return False
'            End If

'            '�T�u�t�H���_����������i�w�肵���t�H���_�j
'            For Each strDirPath As String In System.IO.Directory.GetDirectories(objZipDirectory.FileName)
'                Dim objZipSubDirectory As ClsZipFile = New ClsZipFile()
'                objZipSubDirectory.FileName = strDirPath & "\"
'                objZipSubDirectory.FileTimestamp = System.IO.File.GetLastWriteTime(strDirPath)
'                If CompressDirectoryRelativePath(objZipSubDirectory, objZipFileStream, blnIsFolderSave, strTopDirectory) = False Then Return False
'            Next

'            '����
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'#End Region

'#Region "�t�@�C�����k�����i���΃p�X�j�@CompressFileRelativePath"

'    '**********************************************************************
'    ' Name       :�t�@�C�����k�����i���΃p�X�j
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :�w�肳�ꂽ�t�@�C���𑊑΃p�X�ň��k����B
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :objZipFile        -   ���k����t�@�C��
'    '             objZipFileStream  -   ���k�`��
'    '             blnIsFolderSave   -   True    =   �t�H���_���ێ�����ꍇ
'    '                                   False   =   �t�H���_���ێ����Ȃ��ꍇ
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Boolean           -   ���g�p
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2008/08/26
'    '        By  :FMS)K.Kojima
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Private Function CompressFileRelativePath(ByVal objZipFile As ClsZipFile, _
'                                              ByVal objZipFileStream As java.util.zip.ZipOutputStream, _
'                                              ByVal blnIsFolderSave As Boolean, _
'                                     Optional ByVal strTopDirectory As String = "") As Boolean
'        Try
'            Dim strZipFile_FileName As String = objZipFile.FileName
'            Dim strTop_Directory As String = strTopDirectory
'            If strZipFile_FileName.Substring(0, 2) = "\\" Then
'                strZipFile_FileName = "?:" & strZipFile_FileName.Remove(0, strZipFile_FileName.IndexOf("\", 2))
'            End If
'            If strTop_Directory.Substring(0, 2) = "\\" Then
'                strTop_Directory = "?:" & strTop_Directory.Remove(0, strTop_Directory.IndexOf("\", 2))
'            End If

'            ' �t�@�C�����̓X�g���[���̍쐬
'            Dim strFileName As String = strZipFile_FileName
'            If blnIsFolderSave = True Then ' �t�H���_���ێ�����ꍇ
'                ' ���[�g���폜���A\��/�ɒu�����Ă���
'                strFileName = strFileName.Remove(0, System.IO.Path.GetPathRoot(strFileName).Length).Replace("\", "/")
'            Else ' �t�@�C�����݂̂̏ꍇ
'                ' �t�@�C�����{�݂̂̂𒊏o����
'                strFileName = System.IO.Path.GetFileName(strFileName)
'            End If

'            ' �W�j�o�C�g�����̓t�@�C������ǂݍ��݁A�o�̓t�@�C���֏�����
'            Dim maxFileName As String
'            If blnIsFolderSave = True Then ' �t�H���_���ێ�����ꍇ
'                If strZipFile_FileName.ToUpper = strTop_Directory.ToUpper Then
'                    maxFileName = strFileName & "/"
'                Else
'                    maxFileName = IO.Path.GetDirectoryName(strTop_Directory)
'                    If maxFileName.ToUpper = System.IO.Path.GetPathRoot(strTop_Directory).ToUpper Then
'                        maxFileName = strZipFile_FileName.Remove(0, maxFileName.Length).Replace("\", "/")
'                    Else
'                        maxFileName = strZipFile_FileName.Remove(0, maxFileName.Length + 1).Replace("\", "/")
'                    End If
'                    If strZipFile_FileName.Substring(strZipFile_FileName.Length - 1, 1) = "\" Then
'                        If maxFileName.Substring(maxFileName.Length - 1, 1) <> "/" Then maxFileName &= "/"
'                    End If
'                End If
'            Else
'                maxFileName = strFileName
'            End If
'            'maxFileName = Mid(strFileName, strFileName.IndexOf("GTX/") + 1)
'            Dim objZipEntry As New java.util.zip.ZipEntry(maxFileName)
'            objZipEntry.setMethod(java.util.zip.ZipEntry.DEFLATED)

'            '��hayashi 2010/01/26
'            '���k���t�@�C���̍X�V���t���G�|�b�NTime�Őݒ肷��
'            '(�^�C���X�^���v���ݒ肳��Ă��Ȃ��ꍇ�͐ݒ肵�Ȃ�)
'            If objZipFile.FileTimestamp.ToOADate <> 0.0 Then
'                objZipEntry.setTime(toEpochTime(objZipFile.FileTimestamp))
'            End If
'            '��hayashi 2010/01/26

'            objZipFileStream.putNextEntry(objZipEntry)
'            Dim bytBuffer(8191) As System.SByte
'            Dim lngLength As Long = 0
'            Dim intCompressedLength As Integer = 0
'            Dim intReadedLength As Integer
'            Dim objFileInputStream As java.io.FileInputStream = Nothing

'            '0529
'            Dim crc As New java.util.zip.CRC32

'            If Not IO.Path.GetFileName(objZipFile.FileName) = String.Empty Then
'                objFileInputStream = New java.io.FileInputStream(objZipFile.FileName)

'                Do While True
'                    intReadedLength = objFileInputStream.read(bytBuffer, 0, bytBuffer.Length)
'                    If intReadedLength <= 0 Then
'                        Exit Do
'                    End If

'                    '0529
'                    crc.update(bytBuffer, 0, intReadedLength)

'                    objZipFileStream.write(bytBuffer, 0, intReadedLength)
'                    lngLength += intReadedLength
'                Loop

'            End If

'            bytBuffer = New System.SByte() {}
'            bytBuffer = Nothing

'            '2008/05/22Dim lngTime As Long = java.util.Date.parse(String.Format("{0:yyyy/MM/dd HH:mm:ss}", System.IO.File.GetLastWriteTime(objZipFile.FileName)))
'            '2008/05/22objZipEntry.setTime(lngTime)
'            '2008/05/28objZipEntry.setSize(lngLength)
'            objZipEntry.setSize(lngLength)
'            objZipEntry.setCrc(crc.getValue())

'            ' ���̓X�g���[�������
'            If Not objFileInputStream Is Nothing Then
'                objFileInputStream.close()
'                objFileInputStream = Nothing
'            End If

'            ' �y�h�o�G���g����ǉ�����
'            objZipFileStream.closeEntry()

'            '0529
'            objZipFileStream.flush()

'            crc = Nothing

'            '����
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'#End Region
'    '�� Add by K.Kojima 2008/08/26 End  �i�P�DFrmRestoreTool.exe�����s����t�H���_���AGTX�t�H���_����GTX_Data�ɍ쐬�����e���|�����t�H���_�Ŏ��s����l�ɕύX�j*************************************************************************

'    '**********************************************************************
'    ' Name       :���t�^���G�|�b�N�^�C���ɕϊ�����
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :dt -  ���t�^
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Long - �G�|�b�N�^�C��
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2010/01/26
'    '        By  :hayashi
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Private Shared Function toEpochTime(ByVal dt As DateTime) As Long
'        Dim lngEpo As Long = CType((dt.ToUniversalTime().Ticks - 621355968000000000) / 10000, Long)
'        Return lngEpo
'    End Function

'    '**********************************************************************
'    ' Name       :�G�|�b�N�^�C������t�^�ɕϊ�����
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :dt -  �G�|�b�N�^�C��
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Long - ���t�^
'    '
'    '----------------------------------------------------------------------
'    ' Design On  :2010/01/26
'    '        By  :hayashi
'    '----------------------------------------------------------------------
'    ' Update On  :
'    '        By  :
'    ' Comment    :
'    '**********************************************************************
'    Private Shared Function fromEpochTime(ByVal lngEpo As Long) As DateTime
'        Dim dt As DateTime = New DateTime(lngEpo * 10000 + 621355968000000000)
'        Return dt.ToLocalTime()
'    End Function


'End Class
