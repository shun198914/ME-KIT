'Option Strict On
''**************************************************************************
'' Name       :ＺＩＰファイル制御クラスtestx
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
'    ' バブリックイベント
'    '**********************************************************************
'    Public Event OperationInProgress(ByVal isTopItem As Boolean, ByVal TagetPath As String, ByVal Now As Integer, ByVal total As Integer, ByRef Cancel As Boolean)

'    Public Class ClsZipFile
'        '**********************************************************************
'        ' プライベート変数
'        '**********************************************************************
'        Private m_strFileName As String ' ファイル名
'        Private m_lngFileSize As Long ' ファイルサイズ
'        Private m_lngCompressedSize As Long ' 圧縮ファイルサイズ
'        Private m_dteTimeStamp As Date ' タイムスタンプ

'        '**********************************************************************
'        ' プライベート変数
'        '**********************************************************************

'        '**********************************************************************
'        ' Name       :ファイル名プロパティ
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
'        ' Name       :ファイルサイズプロパティ
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
'        ' Name       :圧縮ファイルサイズプロパティ
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
'        ' Name       :ファイルタイムスタンププロパティ
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
'        ' プライベート変数
'        '**********************************************************************
'        Private m_objZipFiles As Hashtable

'        '**********************************************************************
'        ' Name       :コンストラクタ
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
'        ' Name       :ファイル個数
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
'        ' Name       :圧縮ファイルコレクションの登録
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
'        ' Name       :圧縮ファイルコレクションの１エントリ
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
'    ' Name       :圧縮ファイル一覧をリスト出力する
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :2010/01/26現在、誰も使用していない（バグ有り）の為、
'    '             使用禁止
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

'            '開くZIPファイルの設定
'            Dim zipPath As String = strZipPath

'            '読み込む
'            Dim objFileStream As New java.io.FileInputStream(zipPath)
'            Dim objZipInputStream As New java.util.zip.ZipInputStream(objFileStream)

'            'ZIP内のファイル情報を取得
'            Do While True
'                Dim objZipEntry As java.util.zip.ZipEntry = objZipInputStream.getNextEntry()
'                If (objZipEntry Is Nothing) Then
'                    Exit Do
'                End If
'                If Not objZipEntry.isDirectory() Then
'                    '情報を表示する
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

'            '閉じる
'            objZipInputStream.close()
'            objFileStream.close()

'            Return objZipFiles

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :指定されたファイルを圧縮する
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
'            ' 基本の圧縮処理呼び出し用クラスにファイル名をセットする
'            Dim intCompressedFileCount As Integer = 0
'            Dim objTargetZipFiles As ClsZipFiles
'            objTargetZipFiles = New ClsZipFiles()
'            For intFileIndex As Integer = 0 To objZipFiles.Count - 1
'                Dim objTargetZipFile As ClsZipFile = New ClsZipFile()
'                objTargetZipFile.FileName = CType(objZipFiles(intFileIndex), String)
'                objTargetZipFiles.Add(objTargetZipFile)
'            Next

'            ' 基本の圧縮処理を呼び出す
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
'            ' ＺＩＰ出力ストリームの作成
'            objFileStream = New java.io.FileOutputStream(strZipFileName)
'            objZipFileStream = New java.util.zip.ZipOutputStream(objFileStream)

'            Dim timer As New DateTime

'            ' すべてのファイルに対して処理する
'            For intRecordIndex As Integer = 0 To objZipFiles.Count - 1

'                If timer.Second <> Now.Second _
'                OrElse intCompressedFileCount = 0 _
'                OrElse intCompressedFileCount = (objZipFiles.Count - 1) Then
'                    'イベントを起こします。
'                    RaiseEvent OperationInProgress(True, objZipFiles.Item(intRecordIndex).FileName, intCompressedFileCount + 1, objZipFiles.Count, blnCancel)

'                    'キャンセルされた場合は繰り返しを抜けます。
'                    If blnCancel = True Then Return -1

'                    timer = Now
'                End If

'                If System.IO.File.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' ファイルの場合
'                    If CompressFile(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                ElseIf System.IO.Directory.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' ディレクトリの場合
'                    If CompressDirectory(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                End If

'                ' 圧縮ファイル数の更新
'                intCompressedFileCount += 1
'            Next

'            Return intCompressedFileCount


'        Catch ex As Exception
'            Throw
'        Finally
'            ' ＺＩＰ出力ストリームを閉じる
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
'    ' Comment    :なんかこのメソッドはダメだと思います。
'    '**********************************************************************
'    Private Function CompressDirectory(ByVal objZipDirectory As ClsZipFile _
'                                     , ByVal objZipFileStream As java.util.zip.ZipOutputStream _
'                                     , ByVal blnIsFolderSave As Boolean) As Boolean
'        Try
'            Dim strFiles As String() = System.IO.Directory.GetFiles(objZipDirectory.FileName)

'            For Each strFilePath As String In strFiles
'                ' ファイルを圧縮する
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = strFilePath
'                CompressFile(objZipFile, objZipFileStream, blnIsFolderSave)
'            Next

'            If strFiles.Length = 0 Then
'                Dim objZipFile As ClsZipFile = New ClsZipFile()
'                objZipFile.FileName = objZipDirectory.FileName
'                CompressFile(objZipFile, objZipFileStream, blnIsFolderSave)
'            End If

'            'サブフォルダを処理する（指定したフォルダ）
'            For Each strDirPath As String In System.IO.Directory.GetDirectories(objZipDirectory.FileName)
'                Dim objZipSubDirectory As ClsZipFile = New ClsZipFile()
'                objZipSubDirectory.FileName = strDirPath & "\"
'                CompressDirectoryRelativePath(objZipSubDirectory, objZipFileStream, blnIsFolderSave)
'            Next

'            '正常
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :ファイル圧縮
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
'    ' Comment    :解凍後の日付が異常に古くなる現象があり調査したところ
'    '             下記でエントリの修正時間を設定しているが、不要と思われたので
'    '             コメントしたら正常な日付となった
'    ' Update On  :2008/05/28
'    '        By  :FMS)N.Kasama
'    ' Comment    :「ローカルヘッダーのデータが不正です」というメッセージ対応
'    '**********************************************************************
'    Private Function CompressFile(ByVal objZipFile As ClsZipFile _
'                                , ByVal objZipFileStream As java.util.zip.ZipOutputStream _
'                                , ByVal blnIsFolderSave As Boolean) As Boolean
'        Try
'            ' ファイル入力ストリームの作成
'            Dim strFileName As String = objZipFile.FileName
'            If blnIsFolderSave = True Then ' フォルダを維持する場合
'                ' ルートを削除し、\を/に置換しておく
'                strFileName = strFileName.Remove(0, System.IO.Path.GetPathRoot(strFileName).Length).Replace("\", "/")
'            Else ' ファイル名のみの場合
'                ' ファイル名本体のみを抽出する
'                strFileName = System.IO.Path.GetFileName(strFileName)
'            End If

'            ' ８Ｋバイトずつ入力ファイルから読み込み、出力ファイルへ書込む
'            Dim objZipEntry As New java.util.zip.ZipEntry(strFileName)
'            objZipEntry.setMethod(java.util.zip.ZipEntry.DEFLATED)

'            '↓hayashi 2010/01/26
'            '圧縮元ファイルの更新日付をエポックTimeで設定する
'            '(タイムスタンプが設定されていない場合は設定しない)
'            If objZipFile.FileTimestamp.ToOADate <> 0.0 Then
'                objZipEntry.setTime(toEpochTime(objZipFile.FileTimestamp))
'                '↑hayashi 2010/01/26
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

'            ' 入力ストリームを閉じる
'            objFileInputStream.close()

'            ' ＺＩＰエントリを追加する
'            objZipFileStream.closeEntry()

'            '0529
'            objZipFileStream.flush()

'            '正常
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '**********************************************************************
'    ' Name       :指定された圧縮ファイルをすべて展開する
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
'            '読み込む
'            Dim objFileStream As New java.io.FileInputStream(strSrcZipFileName)
'            Dim objZipStream As New java.util.zip.ZipInputStream(objFileStream)
'            If System.IO.Directory.Exists(strDestExpandFolderName) = False Then
'                Try
'                    System.IO.Directory.CreateDirectory(strDestExpandFolderName)
'                Catch
'                    Throw New Exception("以下のフォルダを作成できませんでした。" & ControlChars.CrLf & strDestExpandFolderName)
'                End Try
'            End If

'            Dim timer As New DateTime
'            Dim blnCancel As Boolean = False
'            Dim intI As Integer = 0
'            Dim dtFileTimestamp As Date
'            Dim strFileName As String = ""

'            'ZIP内のファイル情報を取得
'            Dim objZipEntry As java.util.zip.ZipEntry = objZipStream.getNextEntry()

'            Do Until objZipEntry Is Nothing

'                If Not objZipEntry.isDirectory() Then

'                    'ファイル名
'                    strFileName = objZipEntry.getName()

'                    If timer.Second <> Now.Second _
'                    OrElse intI = 0 Then
'                        'イベントを起こします。
'                        RaiseEvent OperationInProgress(True, strFileName, intI + 1, -1, blnCancel)

'                        'キャンセルされた場合は繰り返しを抜けます。
'                        If blnCancel = True Then Return -1

'                        timer = Now
'                    End If

'                    intI += 1

'                    '展開先のパス
'                    Dim strOutputFolderName As String = System.IO.Path.Combine(strDestExpandFolderName, System.IO.Path.GetDirectoryName(strFileName))
'                    If System.IO.Directory.Exists(strOutputFolderName) = False Then ' フォルダが存在しない場合
'                        ' フォルダを作成する
'                        System.IO.Directory.CreateDirectory(strOutputFolderName)
'                    End If
'                    Dim strOutputFileName As String = System.IO.Path.Combine(strOutputFolderName, System.IO.Path.GetFileName(strFileName))
'                    'FileOutputStreamの作成
'                    Dim objFileOutputStream As java.io.FileOutputStream
'                    Try
'                        objFileOutputStream = New java.io.FileOutputStream(strOutputFileName)
'                    Catch
'                        Throw New Exception("以下のファイルを作成できませんでした。" & ControlChars.CrLf & strOutputFileName _
'                                           & (intI - 1).ToString & "個のファイルは展開されています。")
'                    End Try
'                    '書込み
'                    Dim bytBuffer(8191) As System.SByte

'                    Do While True
'                        Dim len As Integer = objZipStream.read(bytBuffer, 0, bytBuffer.Length)
'                        If len <= 0 Then
'                            Exit Do
'                        End If
'                        objFileOutputStream.write(bytBuffer, 0, len)
'                    Loop

'                    ' 出力ストリームを閉じる
'                    objFileOutputStream.close()
'                    objFileOutputStream.Finalize()
'                    objFileOutputStream = Nothing

'                    bytBuffer = New System.SByte() {}
'                    bytBuffer = Nothing

'                    'エポックTimeを変換して更新日を格納
'                    dtFileTimestamp = fromEpochTime(objZipEntry.getTime())

'                    ' タイムスタンプを圧縮ファイルの情報の置換する
'                    System.IO.File.SetCreationTime(strOutputFileName, dtFileTimestamp)

'                    System.IO.File.SetLastWriteTime(strOutputFileName, dtFileTimestamp)

'                    System.IO.File.SetLastAccessTime(strOutputFileName, dtFileTimestamp)
'                End If

'                objZipStream.closeEntry()

'                'ZIP内のファイル情報を取得
'                objZipEntry = objZipStream.getNextEntry()
'            Loop

'            If intI > 0 Then
'                'イベントを起こします。
'                RaiseEvent OperationInProgress(True, strFileName, intI + 1, -1, blnCancel)
'            End If

'            ' 入力ストリームを閉じる()
'            objZipStream.close()
'            objFileStream.close()

'            Return intI

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'    '↓ Add by K.Kojima 2008/08/26 Start（１．FrmRestoreTool.exeを実行するフォルダを、GTXフォルダからGTX_Dataに作成したテンポラリフォルダで実行する様に変更）*************************************************************************
'#Region "指定ディレクトリ圧縮処理（相対パス）　CompressZipFilesRelativePath"

'    '**********************************************************************
'    ' Name       :指定ディレクトリ圧縮処理（相対パス）
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :指定されたディレクトリを相対パスで圧縮処理を行う
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :strZipFileName    -   圧縮したZIPファイルパス名（フルパス）
'    '             objZipFiles       -   圧縮するディレクトリ
'    '             blnIsFolderSave   -   True    =   フォルダを維持する場合
'    '                                   False   =   フォルダを維持しない場合
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Integer           -   圧縮ファイル数
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
'                ' ＺＩＰ出力ストリームの作成
'                objFileStream = New java.io.FileOutputStream(strZipFileName)
'                objZipFileStream = New java.util.zip.ZipOutputStream(objFileStream)

'                Dim timer As New DateTime

'                ' すべてのファイルに対して処理する
'                For intRecordIndex As Integer = 0 To objZipFiles.Count - 1

'                    If timer.Second <> Now.Second _
'                    OrElse intCompressedFileCount = 0 _
'                    OrElse intCompressedFileCount = (objZipFiles.Count - 1) Then
'                        'イベントを起こします。
'                        RaiseEvent OperationInProgress(True, objZipFiles.Item(intRecordIndex).FileName, intCompressedFileCount + 1, objZipFiles.Count, blnCancel)

'                        'キャンセルされた場合は繰り返しを抜けます。
'                        If blnCancel = True Then Return -1

'                        timer = Now
'                    End If

'                    If System.IO.File.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' ファイルの場合
'                        If CompressFileRelativePath(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave) = False Then Return -1
'                    ElseIf System.IO.Directory.Exists(objZipFiles.Item(intRecordIndex).FileName) = True Then ' ディレクトリの場合
'                        If CompressDirectoryRelativePath(objZipFiles.Item(intRecordIndex), objZipFileStream, blnIsFolderSave, objZipFiles.Item(intRecordIndex).FileName) = False Then Return -1
'                    End If

'                    ' 圧縮ファイル数の更新
'                    intCompressedFileCount += 1
'                Next

'                '正常
'                Return intCompressedFileCount

'            Catch ex As Exception
'                Throw
'            Finally
'                ' ＺＩＰ出力ストリームを閉じる
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

'#Region "ディレクトリ圧縮処理（相対パス）　CompressDirectoryRelativePath"

'    '**********************************************************************
'    ' Name       :ディレクトリ圧縮処理（相対パス）
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :指定されたディレクトリに存在するファイルを圧縮し、
'    '             フォルダが存在した場合は再帰し、圧縮処理を行う
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :objZipDirectory   -   圧縮するディレクトリ
'    '             objZipFileStream  -   圧縮形式
'    '             blnIsFolderSave   -   True    =   フォルダを維持する場合
'    '                                   False   =   フォルダを維持しない場合
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Boolean           -   未使用
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
'                    'イベントを起こします。
'                    RaiseEvent OperationInProgress(False, objZipDirectory.FileName, intI + 1, strFiles.Length, blnCancel)

'                    'キャンセルされた場合は繰り返しを抜けます。
'                    If blnCancel = True Then Return False

'                    timer = Now
'                End If

'                ' ファイルを圧縮する
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

'            'サブフォルダを処理する（指定したフォルダ）
'            For Each strDirPath As String In System.IO.Directory.GetDirectories(objZipDirectory.FileName)
'                Dim objZipSubDirectory As ClsZipFile = New ClsZipFile()
'                objZipSubDirectory.FileName = strDirPath & "\"
'                objZipSubDirectory.FileTimestamp = System.IO.File.GetLastWriteTime(strDirPath)
'                If CompressDirectoryRelativePath(objZipSubDirectory, objZipFileStream, blnIsFolderSave, strTopDirectory) = False Then Return False
'            Next

'            '正常
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'#End Region

'#Region "ファイル圧縮処理（相対パス）　CompressFileRelativePath"

'    '**********************************************************************
'    ' Name       :ファイル圧縮処理（相対パス）
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :指定されたファイルを相対パスで圧縮する。
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :objZipFile        -   圧縮するファイル
'    '             objZipFileStream  -   圧縮形式
'    '             blnIsFolderSave   -   True    =   フォルダを維持する場合
'    '                                   False   =   フォルダを維持しない場合
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Boolean           -   未使用
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

'            ' ファイル入力ストリームの作成
'            Dim strFileName As String = strZipFile_FileName
'            If blnIsFolderSave = True Then ' フォルダを維持する場合
'                ' ルートを削除し、\を/に置換しておく
'                strFileName = strFileName.Remove(0, System.IO.Path.GetPathRoot(strFileName).Length).Replace("\", "/")
'            Else ' ファイル名のみの場合
'                ' ファイル名本体のみを抽出する
'                strFileName = System.IO.Path.GetFileName(strFileName)
'            End If

'            ' ８Ｋバイトずつ入力ファイルから読み込み、出力ファイルへ書込む
'            Dim maxFileName As String
'            If blnIsFolderSave = True Then ' フォルダを維持する場合
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

'            '↓hayashi 2010/01/26
'            '圧縮元ファイルの更新日付をエポックTimeで設定する
'            '(タイムスタンプが設定されていない場合は設定しない)
'            If objZipFile.FileTimestamp.ToOADate <> 0.0 Then
'                objZipEntry.setTime(toEpochTime(objZipFile.FileTimestamp))
'            End If
'            '↑hayashi 2010/01/26

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

'            ' 入力ストリームを閉じる
'            If Not objFileInputStream Is Nothing Then
'                objFileInputStream.close()
'                objFileInputStream = Nothing
'            End If

'            ' ＺＩＰエントリを追加する
'            objZipFileStream.closeEntry()

'            '0529
'            objZipFileStream.flush()

'            crc = Nothing

'            '正常
'            Return True

'        Catch ex As Exception
'            Throw
'        End Try
'    End Function

'#End Region
'    '↑ Add by K.Kojima 2008/08/26 End  （１．FrmRestoreTool.exeを実行するフォルダを、GTXフォルダからGTX_Dataに作成したテンポラリフォルダで実行する様に変更）*************************************************************************

'    '**********************************************************************
'    ' Name       :日付型をエポックタイムに変換する
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :dt -  日付型
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Long - エポックタイム
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
'    ' Name       :エポックタイムを日付型に変換する
'    '
'    '----------------------------------------------------------------------
'    ' Comment    :
'    '
'    '----------------------------------------------------------------------
'    ' Parameter  :dt -  エポックタイム
'    '
'    '----------------------------------------------------------------------
'    ' Return     :Long - 日付型
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
