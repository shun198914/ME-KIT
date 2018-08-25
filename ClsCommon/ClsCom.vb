Option Strict On
'**************************************************************************
' Name       :ClsCom 共通ライブラリ
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
    ' クラス固有メンバー
    '**********************************************************************
    Private Shared m_recXSys As Rec_XSys
    '**********************************************************************
    'イベント
    '**********************************************************************
    Public Shared Event CmnLogOut(ByVal ex As Exception _
                                , ByVal enmLogType As ClsCst.enmLogType _
                                , ByVal strText As String _
                                , ByVal blnDebugLogOutOnly As Boolean)

    '**********************************************************************
    ' Name       :ログ/例外出力
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
    ' Name       :XMLファイルデータ取得  
    '
    '----------------------------------------------------------------------
    ' Comment    :strFileName  'XMLファイル名
    '            :strSection   'セクション名
    '            :strKey 　　  'キー名
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
            '戻り値の初期化
            Dim strReturn As String = ""

            ''XMLファイルを読み込む。
            Dim objXmlDocument As New System.Xml.XmlDocument
            objXmlDocument.Load(strFileName)

            '指定セクションのノードリストを取得
            Dim objNodeList As System.Xml.XmlNodeList = objXmlDocument.GetElementsByTagName(strSection)
            If Not objNodeList Is Nothing Then
                '同じセクション名があった場合先頭のノードを使用する
                Dim childNodeList As System.Xml.XmlNodeList = objNodeList.Item(0).ChildNodes
                If Not childNodeList Is Nothing Then
                    'セクションの中のノード（指定キー）を検索
                    For Each objNode As System.Xml.XmlNode In childNodeList
                        If objNode.Name = strKey Then
                            '設定値取得し戻り値として返す
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
    ' Name       :システム設定情報
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
    ' Name       :アプリケーションパス
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
                    ''OutPutLog("の上位に" & ClsCommon.ClsCst.FldrPath.System & "フォルダが見つかりませんでした。")
                End If
                'システムパスを返します。
                Return GetAppPath

            Catch ex As Exception
                Throw
            End Try
        End Get
    End Property

    '**********************************************************************
    ' Name       :イメージパス
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '　　　　　　　
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On　: 
    '        By　: 
    ' Comment  　: 
    '**********************************************************************
    Public Shared ReadOnly Property GetImagePath() As String
        Get
            '画像パスを返します。
            Return GetAppPath & ClsCst.FldrName.Image & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :サウンドパス
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '　　　　　　　
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On　: 
    '        By　: 
    ' Comment  　: 
    '**********************************************************************
    Public Shared ReadOnly Property GetSoundPath() As String
        Get
            '画像パスを返します。
            Return GetAppPath & ClsCst.FldrName.Sound & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :ログファイルパス
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '　　　　　　　
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
            'ログファイルパスを返します。
            Return GetAppPath & ClsCst.FldrName.Log & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :SQLスクリプトファイルパス
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '　　　　　　　
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
            'SQLスクリプトファイルパスを返します。
            Return GetAppPath & ClsCst.FldrName.SqlScript & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :ヘルプファイルパス
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '　　　　　　　
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yoshida
    '        By  :2007/12/01
    '----------------------------------------------------------------------
    ' Update On  : 
    '        By　: 
    ' Comment    : 
    '**********************************************************************
    Public Shared ReadOnly Property GetHelpPath() As String
        Get
            '画像パスを返します。
            Return GetAppPath & ClsCst.FldrName.Help & "\"
        End Get
    End Property

    '**********************************************************************
    ' Name       :文字列を引用符で挟む
    '
    '----------------------------------------------------------------------
    ' Comment    :strText  文字列
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
            '元々の文字列が引用符を含む場合の変換が必要
            Return "'" & strText.Replace("'", "''") & "'"
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :ダブルクォート編集（オート）
    '
    '----------------------------------------------------------------------
    ' Comment    :文字型、キャラクタ型、日付型は、ダブルクォート編集後、前後"付加
    '　　　　　  　 それ以外は単純にString変換して返します。
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
                    '元々の文字列が引用符を含む場合の変換が必要
                    Return "'" & CType(objMember, String).Replace("'", "''") & "'"
                Case Else
                    'そのまま文字列化して返す
                    Return CType(objMember, String)
            End Select

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :文字列を改行で連結します。
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
            'CrLfで行を結合します。
            strBase = strBase & ControlChars.CrLf & strAddText
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :文字列を改行で連結します。
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
            'CrLfで行を結合します。
            strBase.Append(ControlChars.CrLf & strAddText)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :文字列をカンマ文字付き改行で連結します。
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

            'CrLfで行を結合します。
            strBase.Append(ControlChars.CrLf & strJoint & strAddText)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :文字列をWHERE又はAND文字付き改行で連結します。
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

            'CrLfで行を結合します。
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :文字列をORDER BY又はカンマ文字付き改行で連結します。
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

            'CrLfで行を結合します。
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :文字列をSET又はカンマ文字付き改行で連結します。
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

            'CrLfで行を結合します。
            strBase.Append(ControlChars.CrLf & strJoint)
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ファイル／フォルダパス名の代替文字列変換
    '
    '----------------------------------------------------------------------
    ' Comment    :%AppRoot% →FrmMenu.Exeのあるフォルダパスにします。
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
            '以下に象徴文字列分のリプレース文を書いて下さい。
            CnvPath = Replace(strPath, ClsCst.FldrName.AppRoot, GetAppPath, 1, -1, CompareMethod.Text)
            'CnvPath = strPath.Replace(ClsCst.FldrName.AppRoot, GetAppPath)

            Return CnvPath

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :文字列の暗号化
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

            Dim objCryptoServiceProvider As New System.Security.Cryptography.RC2CryptoServiceProvider '暗号化を実施するオブジェクト
            Dim bytInPutWords As Byte() = System.Text.Encoding.UTF8.GetBytes(pKeyWord)                '暗号化対象バイト列（パラメタをバイト列にエンコード）
            Dim bytOutPutWords As Byte()                                                              '暗号化後のバイト列
            Dim objOutPutSteam As New System.IO.MemoryStream                                          '暗号化データを書き込むオブジェクト

            '暗号化鍵設定
            objCryptoServiceProvider.Key = System.Text.Encoding.UTF8.GetBytes("YkAeMnAjSiH$I\TA")

            '初期化ベクタ設定
            objCryptoServiceProvider.IV = System.Text.Encoding.UTF8.GetBytes("AjiDAS-6")

            'CryptoStreamオブジェクトの作成
            Dim objTransform As System.Security.Cryptography.ICryptoTransform = objCryptoServiceProvider.CreateEncryptor
            Dim objCryptoStream As New System.Security.Cryptography.CryptoStream(objOutPutSteam, objTransform, System.Security.Cryptography.CryptoStreamMode.Write)

            '暗号化してStreamに出力
            objCryptoStream.Write(bytInPutWords, 0, bytInPutWords.Length)
            objCryptoStream.FlushFinalBlock()

            'Streamから暗号化データを取得
            bytOutPutWords = objOutPutSteam.ToArray

            'オブジェクトを閉じる
            objCryptoStream.Close()
            objOutPutSteam.Close()

            '復帰値を設定する
            EncryptByRC2 = System.Convert.ToBase64String(bytOutPutWords)
        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :文字列の複合化
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

            Dim objCryptoServiceProvider As New System.Security.Cryptography.RC2CryptoServiceProvider '複号化を実施するオブジェクト
            Dim bytInPutWords As Byte() = System.Convert.FromBase64String(pKeyWord)                   '複号化対象バイト列（パラメタをBASE64でデコード）
            Dim strOutPutWords As String                                                              '複号化後の文字列
            Dim objInPutSteam As New System.IO.MemoryStream(bytInPutWords)                            '複号化データを読み込むオブジェクト

            '暗号化鍵設定
            objCryptoServiceProvider.Key = System.Text.Encoding.UTF8.GetBytes("YkAeMnAjSiH$I\TA")

            '初期化ベクタ設定
            objCryptoServiceProvider.IV = System.Text.Encoding.UTF8.GetBytes("AjiDAS-6")

            'CryptoStreamオブジェクトの作成
            Dim objTransform As System.Security.Cryptography.ICryptoTransform = objCryptoServiceProvider.CreateDecryptor
            Dim objCryptoStream As New System.Security.Cryptography.CryptoStream(objInPutSteam, objTransform, System.Security.Cryptography.CryptoStreamMode.Read)

            '複号化してStreamから入力
            Dim objStreamReader As New System.IO.StreamReader(objCryptoStream, System.Text.Encoding.UTF8)
            strOutPutWords = objStreamReader.ReadToEnd

            'オブジェクトを閉じる
            objStreamReader.Close()
            objCryptoStream.Close()
            objInPutSteam.Close()

            '復帰値を設定する
            DecryptByRC2 = strOutPutWords
        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :イメージリストへの画像読み込み
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
            '戻り値の初期化
            SetImageList = -1

            Dim strImagePath As String = GetImagePath

            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
            For Each strLogFileName As String In System.IO.Directory.GetFiles(strImagePath, strFileFormat)
                Dim objImg As System.Drawing.Image = System.Drawing.Image.FromFile(strLogFileName)
                iltCtl.ImageSize = New System.Drawing.Size(objImg.Width, objImg.Height)
                iltCtl.Images.Add(System.IO.Path.GetFileName(strLogFileName), objImg)
                '戻り値を最後に追加したアイテムのインデックス値にします。
                SetImageList = iltCtl.Images.Count - 1
            Next
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／


        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :WAVEファイルの再生
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
            '読み込む
            Dim player As New System.Media.SoundPlayer(GetSoundPath & waveFile)

            '非同期再生する
            player.Play()

            'Throw New DivideByZeroException

            'Beep()


            '次のようにすると、ループ再生される
            'player.PlayLooping()

            '次のようにすると、最後まで再生し終えるまで待機する
            'player.PlaySync()
        Catch ex As Exception
            Throw
        End Try
    End Sub


    '**********************************************************************
    ' Name       :選択食区分→選択食名
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
    ' Name       :選択食名→選択食区分
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
    ' Name       :食事時間区分→食事時間名(又は略名)
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
    ' Name       :食事時間名(又は略名)→食事時間区分
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
    ' Name       :文字列内容チェック
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
            'エラーフラグの初期化
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

            'ShiftJisモードの場合
            Dim objEnc As System.Text.Encoding = Nothing
            If ChkInf.SJisMode = True Then
                'エンコードオブジェクトを用意します。
                objEnc = System.Text.Encoding.GetEncoding("shift-jis")
                '文字列をSjis変換した文字にします。
                strText = objEnc.GetString(objEnc.GetBytes(strText))
                '文字列のバイト長を取得します。
                intstrLen = objEnc.GetByteCount(strText)
            End If

            If blnTrimAndPading = True Then
                'トリム処理指定時はトリム後の文字列にします。
                Select Case ChkInf.TrimMode
                    Case Rec_XChkInf.enmTrim.Left : strText = strText.TrimStart
                    Case Rec_XChkInf.enmTrim.Right : strText = strText.TrimEnd
                    Case Rec_XChkInf.enmTrim.LeftRight : strText = strText.Trim
                End Select

                '穴埋め指定時は穴埋め後の文字列にします。
                If ChkInf.TextMaxLen <> 0 AndAlso ChkInf.PadNarrowChar <> "" Then
                    If ChkInf.SJisMode = True Then
                        'ShiftJisモードの場合
                        Select Case ChkInf.PadMode
                            Case Rec_XChkInf.enmPadMode.Left : strText = New String(ChkInf.PadNarrowChar, ChkInf.TextMaxLen - intstrLen) & strText
                            Case Rec_XChkInf.enmPadMode.Right : strText = strText & New String(ChkInf.PadNarrowChar, ChkInf.TextMaxLen - intstrLen)
                        End Select
                    Else
                        'そうでない場合
                        Select Case ChkInf.PadMode
                            Case Rec_XChkInf.enmPadMode.Left : strText = strText.PadLeft(ChkInf.TextMaxLen, ChkInf.PadNarrowChar)
                            Case Rec_XChkInf.enmPadMode.Right : strText = strText.PadRight(ChkInf.TextMaxLen, ChkInf.PadNarrowChar)
                        End Select
                    End If
                End If
            End If

            '文字列の長さを改めて取得します。
            If ChkInf.SJisMode = True Then
                'ShiftJisモードの場合はバイト長を取得します。
                intstrLen = objEnc.GetByteCount(strText)
            Else
                'それ以外は純粋な文字数を取得します。
                intstrLen = strText.Length
            End If

            'エンコードオブジェクトを破棄
            objEnc = Nothing

            '-----------------------------------------------------------------
            '各種チェック
            '-----------------------------------------------------------------

            '必須入力で未入力ならエラー
            If ChkInf.NotEmpty = True AndAlso strText = String.Empty Then
                ChkInf.EmptyErr = True
            End If

            '文字数の最大桁指定がある場合、指定文字数を超えたらエラー
            If ChkInf.TextMaxLen <> 0 AndAlso ChkInf.TextMaxLen < intstrLen Then
                ChkInf.LenTextMaxErr = True
            End If

            '文字数の最小桁指定がある場合、指定文字数以下ならエラー
            If ChkInf.TextMinLen <> 0 AndAlso ChkInf.TextMinLen > intstrLen Then
                ChkInf.LenTextMinErr = True
            End If

            '文字数の最大or最小桁指定がある場合
            If ChkInf.LenTextMaxErr = True OrElse ChkInf.LenTextMinErr = True Then
                ChkInf.LenBetweenErr = True
            End If

            '全角文字禁止で、半角文字以外の文字が含まれている場合はエラー
            If ChkInf.WideCharSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "^[ -~｡-ﾟ\n\f\r\t]*$") = False Then
                ChkInf.WideCharErr = True
            End If

            '改行Tab文字禁止で、改行Tab文字が含まれている場合はエラー
            If ChkInf.CrLfTabSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "[\n\f\r\t]") = True Then
                ChkInf.CrLfTabErr = True
            End If

            '["][,]の入力禁止で、[']["][,]文字が含まれている場合はエラー(※NumericMode=Falseの時のみ有効)
            If ChkInf.NumericMode = False Then
                If ChkInf.FilCtlCharSupport = False AndAlso System.Text.RegularExpressions.Regex.IsMatch(strText, "["",]") = True Then
                    ChkInf.FilCtlCharErr = True
                End If
            End If

            '数値文字指定の場合
            If ChkInf.DecimalOnly = True AndAlso strText.Trim Like New String("#"c, strText.Trim.Length) = False Then
                ChkInf.DecimalErr = True
            End If

            '数値指定の場合
            If ChkInf.NumericMode = True AndAlso strText <> String.Empty Then
                '数値でないならエラー
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
                        '数値の最大／最小値指定がある場合
                        If ChkInf.NumMaxValue <> 0D OrElse ChkInf.NumMinValue <> 0D Then
                            '最小値に満たないならエラー
                            If ChkInf.NumMinValue > CType(strChkNum, Decimal) Then
                                ChkInf.NumericMinErr = True
                            End If
                            '最大値を超えたらエラー
                            If ChkInf.NumMaxValue < CType(strChkNum, Decimal) Then
                                ChkInf.NumericMaxErr = True
                            End If
                            '最大／最小値を超えているならエラー
                            If ChkInf.NumericMinErr = True OrElse ChkInf.NumericMaxErr = True Then
                                ChkInf.NumBetweenErr = True
                            End If
                        End If
                        '小数点の指定桁
                        If ChkInf.DecMaxLen = 0 AndAlso strChkNum.IndexOf("."c) <> -1 Then
                            ChkInf.DecLenErr = True

                        ElseIf ChkInf.DecMaxLen > 0 AndAlso strChkNum.Split("."c).Length = 2 AndAlso strChkNum.Split("."c)(1).Length > ChkInf.DecMaxLen Then
                            ChkInf.DecLenErr = True
                        End If
                    End If
                End If
            End If

            'テキストボックスの編集後文字列全体を返す。
            ChkInf.Text = strText

            'ひとつでもエラーがあれば、Trueを設定します。
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
    ' Name       :Integerの範囲チェック
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
    ' Name       :Stringの範囲チェック
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
    ' Name       :Doubleの範囲チェック
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
    ' Name       :日付の範囲チェック
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
    ' Name       :整数・分母・分子を分数文字列に変換
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
            '確定の場合はアップダウンコンロトールから分数文字列を作成
            Dim strText As String = String.Empty
            If IntNo <> 0 Then
                strText &= IntNo.ToString
            End If
            If Bunbo <> 0 Then
                If strText <> String.Empty Then strText &= "と"
                strText &= Bunshi.ToString & "/" & Bunbo.ToString
            End If
            Return strText.Trim

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :分数文字列から整数・分子・分母のIntegeer()を取得
    '
    '----------------------------------------------------------------------
    ' Comment    :返される配列要素は0:整数、1:分子、2:分母です。
    '             strUseCrがString.Emptyならそれぞれの値は0になります。
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
        Dim intNo As Long = CType(Math.Floor(Value), Integer) '整数部
        If intNo <> 0 Then Value = Value - intNo

        Dim baseValue As Long = 100 '小数部の最大桁数だけ0を追加

        Dim Second As Long = baseValue
        Dim lngValue As Long = CType(Value * Second, Long)
        Dim lngGcm As Long = 0

        'ユークリッドの互除法による精査
        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
        Do While Second > 0
            lngGcm = Second
            Second = lngValue Mod Second
            lngValue = lngGcm
        Loop
        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

        If intNo <> 0 Then
            '整数部がある場合
            CnvFractionString = intNo.ToString
        End If
        If Value <> 0 Then
            '小数部がある場合
            If CnvFractionString <> String.Empty Then CnvFractionString &= "と"
            CnvFractionString &= ((Value * baseValue) / lngGcm).ToString & "/" & (baseValue / lngGcm).ToString
        End If
        Return CnvFractionString

    End Function

    '**********************************************************************
    ' Name       :分数文字列から整数を取得
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
    ' Name       :分数文字列から分子を取得
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
    ' Name       :分数文字列から分母を取得
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
    ' Name       :分数文字列から整数・分子・分母のIntegeer()を取得
    '
    '----------------------------------------------------------------------
    ' Comment    :返される配列要素は0:整数、1:分子、2:分母です。
    '             strUseCrがString.Emptyならそれぞれの値は0になります。
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
                        Return intVal '全て0
                    Else
                        strUseCr = CnvFractionString(CType(strUseCr, Double))
                    End If
                End If

                If IsNumeric(strUseCr) = True Then
                    intVal(0) = CType(strUseCr, Integer)
                Else
                    '分母を設定
                    intVal(2) = CType(strUseCr.Split("/"c)(1), Integer)
                    '整数と分子の文字列だけを取得
                    strUseCr = strUseCr.Split("/"c)(0)
                    If IsNumeric(strUseCr) = True Then
                        intVal(1) = CType(strUseCr, Integer)
                    Else
                        intVal(1) = CType(strUseCr.Split("と"c)(1), Integer)
                        intVal(0) = CType(strUseCr.Split("と"c)(0), Integer)
                    End If
                End If
            End If

            Return intVal

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :端数処理
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....元になる数値 
    '　　　　　　 enmRoundPosition....端数処理する小数点桁位置
    '　　　　　　 enmRoundDiv.........enmRoundPosition位置の小数値に対する処理
    '　　　　　　 例)CnvRound(1.256,enmRoundPosition.P2,enmRoundDiv.Rounded) 
    '                と指定した場合は少数第2位を四捨五入なので結果値は1.3(1.3)になります。
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
                Case ClsCst.enmRoundDiv.Cutoff '切り捨て
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            Return Math.Floor(dblValue * 100000000000000) / 100000000000000
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select

                Case ClsCst.enmRoundDiv.RoundUp '切り上げ
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            Return Math.Round(dblValue + 0.000000000000004, 4, MidpointRounding.AwayFromZero)
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select

                Case ClsCst.enmRoundDiv.Rounded '四捨五入
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            Return Math.Round(dblValue, 14, MidpointRounding.AwayFromZero)
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select
            End Select

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :端数処理
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....元になる数値 
    '　　　　　　 enmRoundPosition....端数処理する小数点桁位置
    '　　　　　　 enmRoundDiv.........enmRoundPosition位置の小数値に対する処理
    '　　　　　　 例)CnvRound(1.256,enmRoundPosition.P2,enmRoundDiv.Rounded) 
    '                と指定した場合は少数第2位を四捨五入なので結果値は1.3(1.3)になります。
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
                Case ClsCst.enmRoundDiv.Cutoff '切り捨て
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            dmlNewValue = Math.Floor(dmlValue * 100000000000000) / 100000000000000
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select

                Case ClsCst.enmRoundDiv.RoundUp '切り上げ
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            dmlNewValue = CType(Math.Round(dmlNewValue + 0.000000000000004, 4, MidpointRounding.AwayFromZero), Decimal)
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select

                Case ClsCst.enmRoundDiv.Rounded '四捨五入
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
                            '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                            dmlNewValue = Math.Round(dmlNewValue, 14, MidpointRounding.AwayFromZero)
                            '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    End Select
            End Select

            Return dmlNewValue

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :切り上げ処理
    '
    '----------------------------------------------------------------------
    ' Comment    :dmlValue.....元になる数値 
    '　　　　　　 enmRoundPosition.........enmRoundPosition位置の小数値に対する処理
    '　　　　　　 例)CnvCutUp(1.250000000003,enmRoundPosition.P3) 
    '                と指定した場合は少数第2位を切り上げなので結果値は(1.26)になります。
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


            'スカラ値関数　[f_get_RoundUP]の型 DECIMAL(13,5)を考慮する
            dmlValue = Math.Floor(dmlValue * 100000) / 100000

            Select Case enmRoundPosition
                Case ClsCst.enmRoundPosition.P1
                    '整数
                    dmlValue = Math.Ceiling(dmlValue)
                Case ClsCst.enmRoundPosition.P2
                    '少数第１位
                    dmlValue = Math.Ceiling(dmlValue * 10) / 10

                Case ClsCst.enmRoundPosition.P3
                    '少数第２位
                    dmlValue = Math.Ceiling(dmlValue * 100) / 100

                Case ClsCst.enmRoundPosition.P4
                    '少数第３位
                    dmlValue = Math.Ceiling(dmlValue * 1000) / 1000

                Case ClsCst.enmRoundPosition.P5
                    '少数第４位
                    dmlValue = Math.Ceiling(dmlValue * 10000) / 10000

                Case ClsCst.enmRoundPosition.P15
                    '少数第１５位
                    '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    dmlValue = Math.Ceiling(dmlValue * 100000000000000) / 100000000000000
                    '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
            End Select

            Return dmlValue
            'Return CType(dblValue, Decimal)

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :切り上げ処理(オーバーロード) :※In Double⇒ Out Decimal
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....元になる数値 (Double)
    '　　　　　　 enmRoundPosition.........enmRoundPosition位置の小数値に対する処理
    '　　　　　　 例)CnvCutUp(1.250000000003,enmRoundPosition.P3) 
    '                と指定した場合は少数第2位を切り上げなので結果値は(1.26)になります。
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
            'スカラ値関数　[f_get_RoundUP]の型 DECIMAL(13,5)を考慮する
            dblValue = Math.Floor(dblValue * 100000) / 100000

            Select Case enmRoundPosition
                Case ClsCst.enmRoundPosition.P1
                    '整数
                    dblValue = Math.Ceiling(dblValue)
                Case ClsCst.enmRoundPosition.P2
                    '少数第１位
                    dblValue = Math.Ceiling(dblValue * 10) / 10

                Case ClsCst.enmRoundPosition.P3
                    '少数第２位
                    dblValue = Math.Ceiling(dblValue * 100) / 100

                Case ClsCst.enmRoundPosition.P4
                    '少数第３位
                    dblValue = Math.Ceiling(dblValue * 1000) / 1000

                Case ClsCst.enmRoundPosition.P5
                    '少数第４位
                    dblValue = Math.Ceiling(dblValue * 10000) / 10000

                Case ClsCst.enmRoundPosition.P15
                    '少数第１５位
                    '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
                    dblValue = Math.Ceiling(dblValue * 100000000000000) / 100000000000000
                    '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
            End Select

            Return CType(dblValue, Decimal)

        Catch ex As Exception
            Throw
        End Try
    End Function


    '**********************************************************************
    ' Name       :端数処理
    '
    '----------------------------------------------------------------------
    ' Comment    :dblValue.....元になる数値 
    '　　　　　　 enmCutUpDiv.........enmCutUpDiv位置の小数値に対する処理
    '　　　　　　 例)CnvCutUp(1.0001,enmCutUpDiv.RoundUp) 
    '                と指定した場合は2.0になります。
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
    '            Case ClsCst.enmCutUpDiv.CutOff  '小数部切捨て
    '                Return Math.Floor(dmlValue)

    '            Case ClsCst.enmCutUpDiv.RoundUp '小数部切り上げ
    '                Return Math.Ceiling(dmlValue)
    '        End Select

    '    Catch ex As Exception
    '        Throw
    '    End Try
    'End Function


    '**********************************************************************
    ' Name       :使用量、廃棄率、換算値から、単位数量を小数点表記の実数に変換します。
    '
    '----------------------------------------------------------------------
    ' Comment    :※blnEasyScrRateには、SysInf.EasyScrRateを設定して下さい。
    '　　　　　　 
    '　　　　　　 
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
    ' Name       :単位数量、廃棄率、換算値から、廃棄率分を差し引いた使用量(g)に変換します。
    '
    '----------------------------------------------------------------------
    ' Comment    :※blnEasyScrRateには、SysInf.EasyScrRateを設定して下さい。
    '　　　　　　 
    '　　　　　　 
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
    ' Name       :使用量、廃棄率から、廃棄率分を含んだ使用量に変換します。
    '
    '----------------------------------------------------------------------
    ' Comment    :※blnEasyScrRateには、SysInf.EasyScrRateを設定して下さい。
    '　　　　　　 
    '　　　　　　 
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
    ' Name       :使用量、廃棄率から、廃棄率分を含んだ使用量に変換します。
    '
    '----------------------------------------------------------------------
    ' Comment    :※blnEasyScrRateには、SysInf.EasyScrRateを設定して下さい。
    '　　　　　　 
    '　　　　　　 
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
    ' Name       :料理食品構成キーの1食品の文字列を返します。
    '
    '----------------------------------------------------------------------
    ' Comment    :食品コード8桁＋使用量6+2桁＋単位計算区分1桁＋整数2桁＋分子2桁＋分母2桁＋真空調理差分日2桁＋"_"
    '　　　　　　 
    '　　　　　　 
    '----------------------------------------------------------------------
    ' Design On  :AST)M.Yosihda
    '        By  :2008/06/02
    '----------------------------------------------------------------------
    ' Update On  :CM)Kuwahara
    '        By  :2011/02/17
    ' Comment    :食品構成キーに真空調理差分日2桁を追加
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

        '料理食品構成キー（食品コード8桁＋使用量4+2桁＋単位計算区分1桁＋整数2桁＋分子2桁＋分母2桁＋真空調理差分日2桁＋"_"）
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
    ' Name       :指定要素が要素配列の中に存在していればTrueを返す。
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '　　　　　　 
    '　　　　　　 
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
    ' Name       :Const情報取得
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

            ''カンマ区切り文字列を配列にSpitする
            Dim strSpit() As String = aStr.Split(CChar(","))
            ''通常の文字列だった場合
            If strSpit.Length = 1 Then
                objCol_GTblCst = Nothing
            Else
                objCol_GTblCst = New Col_GTblCst

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                For i As Integer = 0 To strSpit.Length - 1 Step 2
                    objRec_GTblCst = New Rec_GTblCst
                    objRec_GTblCst.ConstCode = CInt(strSpit(i))
                    objRec_GTblCst.ConstName = strSpit(i + 1)
                    objCol_GTblCst.Add(objRec_GTblCst)
                Next i
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／
            End If

            '----------------------------------------------
            '例外処理
            '----------------------------------------------
        Catch ex As Exception
            objCol_GTblCst = Nothing
            Throw
        End Try

        Return objCol_GTblCst

    End Function

    '**********************************************************************
    ' Name       :指定したフォルダが読めるまで待機する関数
    '
    '----------------------------------------------------------------------
    ' Comment    :指定したフォルダが読めるまで待機する関数
    '             ＲＡＳ接続や遅いＬＡＮの場合、いきなりＣＯＰＹ命令等
    '　　　　　　 実行すると、エラーとなってしまう為、当関数で待機。
    '
    '----------------------------------------------------------------------
    ' Parameter  :FileName: 監視するファイル
    '             Interval: 監視間隔(ms)
    '             Times:    監視繰り返し回数
    '----------------------------------------------------------------------
    ' Return     :ファイルが読めた場合True、時間切れで諦めた場合False
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
                    'エラーは無視
                End Try

                System.Threading.Thread.Sleep(CInt(IntInterval))

                IntTM = IntTM + 1

            Loop

            '時間切れ失敗
            WaitForConnection = False

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :指定したファイルが読めるまで待機する関数
    '
    '----------------------------------------------------------------------
    ' Comment    :指定したファイルが読めるまで待機する関数
    '             ＲＡＳ接続や遅いＬＡＮの場合、いきなりＣＯＰＹ命令等
    '　　　　　　 実行すると、エラーとなってしまう為、当関数で待機。
    '
    '----------------------------------------------------------------------
    ' Parameter  :FileName: 監視するファイル
    '             Interval: 監視間隔(ms)
    '             Times:    監視繰り返し回数
    '----------------------------------------------------------------------
    ' Return     :ファイルが読めた場合True、時間切れで諦めた場合False
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
                    'エラーは無視
                End Try

                System.Threading.Thread.Sleep(CInt(IntInterval))

                IntTM = IntTM + 1

            Loop

            '時間切れ失敗
            WaitFileRead = False

        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :食事時間から、１つ前の時間区分を返す
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
    ' Name       :食事時間から、１つ前の年月日を返す
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

            '時間区分が朝食だったら
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
    ' Name       :食事時間から、１つ後の時間区分を返す
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
    ' Name       :食事時間から、１つ後の年月日を返す
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

            '時間区分が夕間食だったら
            If intEatTmDiv = ClsCst.enmEatTimeDiv.NightDiv Then
                Return DateTime.Parse(Format(CInt(strYMD), "0000/00/00")).AddDays(1).ToString("yyyyMMdd")
            Else
                Return strYMD
            End If

        Catch ex As Exception
            Throw
        End Try
    End Function

    '↓ Add by K.Kojima 2008/07/16 Start（１．終了時にバックアップ機能を動作させる機能を追加）*************************************************************************
    '**********************************************************************
    ' Name       :XMLファイル指定項目取得処理
    '
    '----------------------------------------------------------------------
    ' Comment    :strFileName   -   XMLファイル（フルパス指定）
    '             strSection    -   セクション名
    '             strKey        -   キー名
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

            'XMLファイルを読み込む。
            objXmlDocument.Load(strFileName)

            '指定セクションのノードリストを取得
            Dim objNodeList As Xml.XmlNodeList

            objNodeList = objXmlDocument.GetElementsByTagName(strSection)

            If Not objNodeList Is Nothing Then
                '同じセクション名があった場合先頭のノードを使用する
                Dim childNodeList As Xml.XmlNodeList
                childNodeList = objNodeList.Item(0).ChildNodes
                If Not childNodeList Is Nothing Then
                    'セクションの中のノード（指定キー）を検索
                    Dim objNode As Xml.XmlNode
                    For Each objNode In childNodeList
                        If objNode.Name = strKey Then
                            '設定値取得し戻り値として返す
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
    '↑ Add by K.Kojima 2008/07/16 End  （１．終了時にバックアップ機能を動作させる機能を追加）*************************************************************************


    '**********************************************************************
    ' Name       :ファイルバイナリ比較
    '
    '----------------------------------------------------------------------
    ' Comment    :同じ内容ならTrueを返します。
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
                    '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                    While (fs1.Position < fs1.Length)
                        If fs1.ReadByte() <> fs2.ReadByte() Then Return False
                    End While
                    '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／
                End Using
            End Using

            Return True

        Catch ex As Exception
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Summary    : 年齢計算
    '----------------------------------------------------------------------
    ' Method     : Public Function GetPerAge
    '
    ' Parameter  : ByVal strBirth As String....日付文字列
    '　　　　　　　 Optional ByVal enumDateInterval As DateInterval = DateInterval.Year
    ' Return     :  
    '
    '----------------------------------------------------------------------
    ' Comment    : 引数で渡された誕生日としての日付文字列(連続８桁の数値又は
    '              IsDate関数でTrueとなる日付文字列)とローカルシステム日付を元に
    '              本日時点の年齢値(第２引数によって値は変わります。)を返します。
    '              誕生日付が不正値やNothing、未来日の場合は-1を返します。
    '              ※2/29生まれでシステム年が閏年でない人は3/01生まれとして計算します。【※補足】参照
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
            '誕生日の日付判定
            '---------------------------------------------------------------
            'データが入っていない場合
            If strBirth Is Nothing OrElse strBirth.Trim() = "" Then Return -1
            '８文字の場合は西暦８文字とみなして、日付文字列に変換します。
            If strBirth.Length = 8 AndAlso IsNumeric(strBirth) Then strBirth = CType(strBirth, Integer).ToString("0000/00/00")
            '存在しない日付文字列の場合
            If IsDate(strBirth) = False Then Return -1
            '有効な日付文字列の場合
            Dim dteWork As Date = Date.Parse(strBirth)

            '---------------------------------------------------------------
            '基準日の日付判定
            '---------------------------------------------------------------
            '基準日付
            Dim dteBaseDate As Date
            'データが入っていない場合
            If strBaseDate Is Nothing Then
                dteBaseDate = System.DateTime.Today
            Else
                '８文字の場合は西暦８文字とみなして、日付文字列に変換します。
                If strBaseDate.Length = 8 AndAlso IsNumeric(strBaseDate) Then strBaseDate = CType(strBaseDate, Integer).ToString("0000/00/00")
                '存在しない日付文字列の場合
                If IsDate(strBaseDate) = False Then Return -1
                '有効な日付文字列の場合
                dteBaseDate = Date.Parse(strBaseDate)
            End If

            '基準日付に対して明日以降の誕生日付(未来日)が指定されていた場合は-1を返します。
            If dteWork > dteBaseDate Then Return -1

            If dteWork.ToString("MMdd") = "0229" _
                AndAlso dteBaseDate.Month = 2 _
                AndAlso DateTime.IsLeapYear(dteBaseDate.Year) = False Then
                '2/29生まれの人かつ、基準月が2月かつ、基準年が閏年でない場合は3/1生まれとします。【※補足】参照
                dteWork = dteWork.AddDays(1)
            End If

            '誕生年月日と基準日付とで日付間隔を求めます。
            Dim intRetValue As Long = DateDiff(enumDateInterval, dteWork, dteBaseDate)

            'DateIntervalが年数と月数で指定されている場合には誕生日を迎えたかどうかで日付間隔を補正します。
            Select Case enumDateInterval
                Case DateInterval.Year
                    '基準年の月日が、誕生日日付の月日未満なら、明日以降に誕生日を迎える場合があるので年齢を-1します。
                    If dteBaseDate.ToString("MMdd") < dteWork.ToString("MMdd") Then intRetValue -= 1

                Case DateInterval.Month
                    '基準月の日が誕生日付の日未満で、かつ、明日以降も当月が続くなら、その'明日以降の当月中'に誕生日を迎える場合があるので月数を-1します。
                    If dteBaseDate.ToString("dd") < dteWork.ToString("dd") _
                    AndAlso dteBaseDate.AddDays(1).Month = dteBaseDate.Month Then
                        intRetValue -= 1
                    End If
            End Select

            Return intRetValue

            '┌【※補足】年齢計算ニ関スル法律────────────────────────────────┐
            '│出典: フリー百科事典『ウィキペディア（Wikipedia）』　　　　　　　　　　　　　　　　　　　　　│
            '│年齢計算ニ関スル法律（ねんれいけいさんにかんするほうりつ；明治35年12月2日法律第50号）は、　　│
            '│日本の現行法律の一つであり、年齢の計算方法を定める。民法の附属法の一つに位置づけられる。　　　│
            '│　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　│
            '│本法によれば、年齢は出生の日よりこれを起算し（1項）、暦に従って年数又は月数によって計算する 　│
            '│（2項、民法143条1項。なお、年齢のとなえ方に関する法律1項）。　　　　　　　　　　　　　　　　　│
            '│すなわち、毎年、誕生日（「起算日ニ応答スル日」）の前日をもってある年齢は満了するから　　　　　│
            '│（本法2項、民法143条2項本文）、誕生日の前日が終了する瞬間（誕生日の午前0時00分の直前）に1歳　│
            '│を加えることになる。ただし、誕生日前の最後に始まる月に誕生日が存在しないとき　　　　　　　　　│
            '│（「年ヲ以テ期間ヲ定メタル場合ニ於テ最後ノ月ニ応当日ナキトキ」）は、その月の末日をもってその　│
            '│年齢が満了する（本法2項、民法143条2項但書）。　　　　　　　　　　　　　　　　　　　　　　　　│
            '│　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　│
            '│例えば、2004年10月4日に出生した者の年齢は、誕生日の前日である毎年10月3日が終了する瞬間、　　 │
            '│すなわち10月4日午前0時00分の直前に1歳を加える。　　　　　　　　　　　　　　　　　　　　　　　│
            '│また、2004年（閏年）2月29日に出生した者の年齢は、平年には誕生日前の最後に始まる月である　　　│
            '│2月に誕生日である2月29日が存在しないから、その月の末日である2月28日が終了する瞬間、　　　　　│
            '│すなわち3月1日午前0時00分の直前に、閏年には誕生日の前日である2月28日が終了する瞬間、　　　　 │
            '│すなわち2月29日午前0時00分の直前に、それぞれ1歳を加えるわけである。　　　　　　　　　　　　　│
            '└───────────────────────────────────────────────┘

            '--------------------
        Catch ex As Exception
            Throw
        Finally
        End Try
    End Function

    '**********************************************************************
    ' Name       :文字列の指定バイト分割処理
    '
    '----------------------------------------------------------------------
    ' Comment    :全角を含む文字列を指定バイト以下に分割します。
    '           　例えば、SplitString("123あいう",4)は{"123","あい","う"}を返します。
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

            '文字を一文字づつ走査して20バイト以下にRec分割してColに追加します。
            Dim enc As System.Text.Encoding = System.Text.Encoding.GetEncoding("Shift-JIS")
            Dim intI As Integer = 0
            Dim strExamRst As String = ""
            Dim Col As New Collections.Specialized.StringCollection
            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
            Do
                'intIが文字列の最後か、strExamRstにintI位置の1文字を連結して指定バイト以上になるならRecを追加します
                If intI > (strText.Length - 1) _
                OrElse (enc.GetBytes(strExamRst & strText.Substring(intI, 1)).Length > intSplitLength) Then

                    'Colに追加
                    Col.Add(strExamRst)

                    '次のRecの為にクリア
                    strExamRst = ""
                End If

                'intIが文字数を超えたら繰り返しを抜けます。
                If intI > (strText.Length - 1) Then Exit Do

                'strExamRstに文字位置の示す1文字を連結します。
                strExamRst &= strText.Substring(intI, 1)

                '次の文字位置にカウントアップ
                intI += 1
            Loop
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '文字配列に返します。
            Dim strRetR(Col.Count - 1) As String
            Col.CopyTo(strRetR, 0)

            '返します。
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
    ' Comment    :区切り文字を付加する
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

            ''区切り文字がない場合
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

            '返します。
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

            ''区切り文字がない場合
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
