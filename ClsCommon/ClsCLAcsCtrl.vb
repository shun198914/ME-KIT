Option Strict On
'**************************************************************************
' Name       :クライアントＤＢ制御スーパークラス
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :FTC)S.Shin
'        By  :2011/04/07
'--------------------------------------------------------------------------
' Update On  :
'        By  :
' Comment    :
'**************************************************************************
'**************************************************************************
' Imports
'**************************************************************************
Imports System.Data.SqlClient
Imports ClsCommon
Imports ClsDataFormat
Imports CMN = ClsCommon.ClsCom
Imports CST = ClsCommon.ClsCst

Public Class ClsCLAcsCtrl

    '**********************************************************************
    ' Protected　メンバ変数
    '**********************************************************************
    Public m_objSqlDataSet As DataSet    'データセット
    Public m_intCurrent As Integer = -1  'SQLセレクト後の行位置保持用

    Protected m_objConnection As New SqlClient.SqlConnection    'SqlClient.Sqlコネクション
    Protected m_objTransaction As SqlClient.SqlTransaction      'SqlClientトランザクション
    Protected m_blnTransaction As Boolean                       'トランザクション状態
    Protected m_intCommandTimeout As Integer                    'コマンドタイムアウト

    Private m_recXSys As Rec_XSys                               'システム情報

    '**********************************************************************
    'イベント
    '**********************************************************************
    Public Shared Event CLAcsCtrlLogOut(ByVal ex As Exception _
                                      , ByVal enmLogType As ClsCst.enmLogType _
                                      , ByVal strText As String _
                                      , ByVal blnDebugLogOutOnly As Boolean)

    '**************************************************************************
    ' Name       :初期化
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
    Public Sub New()

        'コモンクラスからのイベントを受け取ります。
        AddHandler ClsCommon.ClsCom.CmnLogOut, AddressOf CmnLogOut

        m_intCommandTimeout = 180
        m_blnTransaction = False

    End Sub

    '**********************************************************************
    ' Name       :トランザクション開始
    '
    '----------------------------------------------------------------------
    ' Comment    :正常終了：0, 異常終了：-1
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub DBBeginTrans()
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            If m_blnTransaction Then     '既にトランザクションが開始されている
                Throw New ApplicationException("既にトランザクションが開始されています。")
            End If
            m_objTransaction = m_objConnection.BeginTransaction()    ''ﾄﾗﾝｻﾞｸｼｮﾝ開始
            m_blnTransaction = True

            Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.Information, "－クライアント側のDBﾄﾗﾝｻﾞｸｼｮﾝ開始－", False)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :DBコミット
    '
    '----------------------------------------------------------------------
    ' Comment    :正常終了：0, 異常終了：-1
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub DBCommit()
        Me.CmnLogOut(Nothing, CST.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            If Not m_blnTransaction Then     'トランザクションが開始されていません。
                Throw New ApplicationException("トランザクションが開始されていません。")
            End If
            m_objTransaction.Commit()
            m_blnTransaction = False
            Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.Information, "－クライアント側のDBコミット実行－", False)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :DBロールバック
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
    Public Sub DBRollBack()
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            m_objTransaction.Rollback()    ''ロールバック
            m_blnTransaction = False
            Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.Information, "－クライアント側のDBロールバック実行－", False)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ExecuteNonQuery
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
    Public Function ExecuteNonQuery(ByVal strCmdText As String, ByVal ColParam As System.Data.SqlClient.SqlParameter()) As Integer
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)

        Try
            Dim objSqlDataAdapter As SqlDataAdapter
            Dim objSqlDbCommand As SqlCommand

            objSqlDataAdapter = New SqlDataAdapter

            If m_blnTransaction Then
                objSqlDbCommand = New SqlCommand(strCmdText, m_objConnection, m_objTransaction)
            Else
                objSqlDbCommand = New SqlCommand(strCmdText, m_objConnection)
            End If

            objSqlDbCommand.CommandTimeout = m_intCommandTimeout         ''タイムアウトを設定

            ''タイプはＳＱＬテキスト
            objSqlDbCommand.CommandType = CommandType.Text

            ''ストアドのパラメーターをセット
            If Not ColParam Is Nothing _
            AndAlso ColParam.Count > 0 Then
                objSqlDbCommand.Parameters.AddRange(ColParam)
            End If

            objSqlDataAdapter.SelectCommand = objSqlDbCommand

            'ExecuteNonQuery では、カタログ操作 (データベース構造の照会、テーブルなどのデータベース オブジェクトの作成など) を実行できます。
            'また、 DataSet を使用せずに、UPDATE、INSERT、または DELETE ステートメントを実行して、データベース内のデータを変更することもできます。
            'ExecuteNonQuery は行(RecordSet)を返しませんが、出力パラメータ、またはパラメータに割り当てられた戻り値には、データが格納されます。
            'UPDATE、INSERT、および DELETE ステートメントでの戻り値は、そのコマンドの影響を受ける行数です。
            '　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '他の種類のステートメントでの戻り値は、-1 です。ロールバックが発生した場合も、戻り値は -1 です。
            Dim lngRecordCount As Integer = objSqlDataAdapter.SelectCommand.ExecuteNonQuery

            Return lngRecordCount

        Catch ex As System.Data.SqlClient.SqlException

            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strCmdText, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)

            Throw
        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :ExecuteScalar
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
    Public Function ExecuteScalar(ByVal strCmdText As String) As Object

        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)

        Try
            Dim objSqlDataAdapter As SqlDataAdapter
            Dim objSqlDbCommand As SqlCommand

            objSqlDataAdapter = New SqlDataAdapter

            If m_blnTransaction Then
                objSqlDbCommand = New SqlCommand(strCmdText, m_objConnection, m_objTransaction)
            Else
                objSqlDbCommand = New SqlCommand(strCmdText, m_objConnection)
            End If

            objSqlDbCommand.CommandTimeout = m_intCommandTimeout        ''タイムアウトを設定

            ''タイプはＳＱＬテキスト
            objSqlDbCommand.CommandType = CommandType.Text


            objSqlDataAdapter.SelectCommand = objSqlDbCommand

            'Transact-SQL ステートメントを実行し、値の型として先頭行の最初の列を返します。
            Dim objRet As Object = objSqlDataAdapter.SelectCommand.ExecuteScalar

            Return objRet

        Catch ex As System.Data.SqlClient.SqlException
            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strCmdText, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :ExcuteInsert
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
    Public Function ExcuteInsert(ByVal strSQL As String) As Integer
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            Dim objSqlDataAdapter As SqlDataAdapter
            Dim objSqlDbCommand As SqlCommand
            Dim lngRecordCount As Integer
            If m_blnTransaction Then
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection, m_objTransaction)
            Else
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection)
            End If

            objSqlDbCommand.CommandTimeout = m_intCommandTimeout        ''タイムアウトを設定

            objSqlDataAdapter = New SqlDataAdapter
            objSqlDataAdapter.InsertCommand = objSqlDbCommand

            'ExecuteNonQuery では、カタログ操作 (データベース構造の照会、テーブルなどのデータベース オブジェクトの作成など) を実行できます。
            'また、 DataSet を使用せずに、UPDATE、INSERT、または DELETE ステートメントを実行して、データベース内のデータを変更することもできます。
            'ExecuteNonQuery は行(RecordSet)を返しませんが、出力パラメータ、またはパラメータに割り当てられた戻り値には、データが格納されます。
            'UPDATE、INSERT、および DELETE ステートメントでの戻り値は、そのコマンドの影響を受ける行数です。
            '　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '他の種類のステートメントでの戻り値は、-1 です。ロールバックが発生した場合も、戻り値は -1 です。

            lngRecordCount = objSqlDataAdapter.InsertCommand.ExecuteNonQuery()

            Return lngRecordCount 'インサート件数を返す。

        Catch ex As System.Data.SqlClient.SqlException
            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strSQL, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :ExcuteUpdate
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
    Public Function ExcuteUpdate(ByVal strSQL As String) As Integer
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            Dim objSqlDataAdapter As SqlDataAdapter
            Dim objSqlDbCommand As SqlCommand
            Dim lngRecordCount As Integer

            If m_blnTransaction Then
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection, m_objTransaction)
            Else
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection)
            End If

            objSqlDbCommand.CommandTimeout = m_intCommandTimeout         ''タイムアウトを設定

            objSqlDataAdapter = New SqlDataAdapter
            objSqlDataAdapter.UpdateCommand = objSqlDbCommand

            'ExecuteNonQuery では、カタログ操作 (データベース構造の照会、テーブルなどのデータベース オブジェクトの作成など) を実行できます。
            'また、 DataSet を使用せずに、UPDATE、INSERT、または DELETE ステートメントを実行して、データベース内のデータを変更することもできます。
            'ExecuteNonQuery は行(RecordSet)を返しませんが、出力パラメータ、またはパラメータに割り当てられた戻り値には、データが格納されます。
            'UPDATE、INSERT、および DELETE ステートメントでの戻り値は、そのコマンドの影響を受ける行数です。
            '　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '他の種類のステートメントでの戻り値は、-1 です。ロールバックが発生した場合も、戻り値は -1 です。
            lngRecordCount = objSqlDataAdapter.UpdateCommand.ExecuteNonQuery()

            Return lngRecordCount '更新件数を返す

        Catch ex As System.Data.SqlClient.SqlException
            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strSQL, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :ExcuteDelete
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
    Public Function ExcuteDelete(ByVal strSQL As String) As Integer

        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)

        Try

            Dim objSqlDataAdapter As SqlDataAdapter
            Dim objSqlDbCommand As SqlCommand
            Dim lngRecordCount As Integer
            If m_blnTransaction Then
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection, m_objTransaction)
            Else
                objSqlDbCommand = New SqlCommand(strSQL, m_objConnection)
            End If

            objSqlDbCommand.CommandTimeout = m_intCommandTimeout         ''タイムアウトを設定

            objSqlDataAdapter = New SqlDataAdapter
            objSqlDataAdapter.DeleteCommand = objSqlDbCommand

            'ExecuteNonQuery では、カタログ操作 (データベース構造の照会、テーブルなどのデータベース オブジェクトの作成など) を実行できます。
            'また、 DataSet を使用せずに、UPDATE、INSERT、または DELETE ステートメントを実行して、データベース内のデータを変更することもできます。
            'ExecuteNonQuery は行(RecordSet)を返しませんが、出力パラメータ、またはパラメータに割り当てられた戻り値には、データが格納されます。
            'UPDATE、INSERT、および DELETE ステートメントでの戻り値は、そのコマンドの影響を受ける行数です。
            '　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
            '他の種類のステートメントでの戻り値は、-1 です。ロールバックが発生した場合も、戻り値は -1 です。

            lngRecordCount = objSqlDataAdapter.DeleteCommand.ExecuteNonQuery()

            Return lngRecordCount '削除件数を返す。

        Catch ex As System.Data.SqlClient.SqlException
            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strSQL, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        End Try
    End Function

    '**********************************************************************
    ' Name       :トランザクション判定
    '
    '----------------------------------------------------------------------
    ' Comment    :Public Function IsOracle() As Boolean
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public ReadOnly Property IsTransaction() As Boolean
        Get
            Return m_blnTransaction
        End Get
    End Property

    '**********************************************************************
    ' Name       :DB接続
    '
    '----------------------------------------------------------------------
    ' Comment    :結果 '正常：0, 異常：-1
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub DBConnect(Optional ByVal ClsCLAcsCtrl As ClsCLAcsCtrl = Nothing)

        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)

        Try
            If ClsCLAcsCtrl Is Nothing Then

                'DB接続文字列の取得
                If SysInf.DBConnectString_CL = "" Then

                    Dim strConnectR(5) As String

                    '各設定値の取得
                    strConnectR(0) = "Data Source    =" & SysInf.DBConnectDataSource_CL
                    strConnectR(1) = "Initial Catalog=" & SysInf.DBConnectInitialCatalog_CL
                    strConnectR(2) = "Connect Timeout=" & SysInf.DBConnectTimeout_CL.ToString
                    strConnectR(3) = "Pooling        =" & SysInf.DBConnectPooling_CL
                    '以下はデコード処理を加えて取得
                    strConnectR(4) = "User ID        =" & SysInf.DBConnectUser_CL
                    strConnectR(5) = "Password       =" & SysInf.DBConnectPassword_CL

                    '一行にします。
                    Me.SysInf.DBConnectString_CL = String.Join(";"c, strConnectR)

                    'SQLコマンドのタイムアウト値を取得
                    m_intCommandTimeout = SysInf.DBConnectCommandTimeout_CL

                End If

                m_objConnection = New SqlClient.SqlConnection(SysInf.DBConnectString_CL)
                m_objConnection.Open()
                Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.Information, "－クライアント側のDB接続開始－", False)

            Else
                'すでにＤＢ接続が確立されている場合は引数で受け取ったDB一連のオブジェクト参照を設定します。
                m_objConnection = ClsCLAcsCtrl.m_objConnection         'SqlClient.Sqlコネクション
                m_objTransaction = ClsCLAcsCtrl.m_objTransaction       'SqlClientトランザクション
                m_blnTransaction = ClsCLAcsCtrl.m_blnTransaction       'トランザクション状態
            End If

        Catch ex As System.Data.SqlClient.SqlException

            Me.CmnLogOut(Nothing, CST.enmLogType.Error, SysInf.DBConnectString_CL, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Catch ex As Exception
            Throw
        End Try
    End Sub

    '**********************************************************************
    ' Name       :DB切断
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
    Public Sub DBDisconnect()

        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)

        Try
            m_objConnection.Close()
            Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.Information, "－クライアント側のDB接続終了－", False)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :Select
    '
    '----------------------------------------------------------------------
    ' Comment    :SQL文を指定し、データを読み込む
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub [Select](ByVal strSelectString As String)
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objSqlDataAdapter As New SqlDataAdapter

            'DataSet取得
            If m_blnTransaction Then
                objSqlDataAdapter.SelectCommand = New SqlCommand(strSelectString, m_objConnection, m_objTransaction)
            Else
                objSqlDataAdapter.SelectCommand = New SqlCommand(strSelectString, m_objConnection)
            End If

            objSqlDataAdapter.SelectCommand.CommandTimeout = m_intCommandTimeout         ''タイムアウトを設定

            m_objSqlDataSet = New DataSet
            objSqlDataAdapter.Fill(m_objSqlDataSet, CST.DataObjName.Table)

            'DataSetが取得できなかった場合、
            'DataSetのTableが存在しない場合、
            'DataSetのTableのrowが存在しない場合はカレントを-1とする
            If m_objSqlDataSet Is Nothing _
            OrElse m_objSqlDataSet.Tables.Count = 0 _
            OrElse m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows.Count = 0 Then
                m_intCurrent = -1
            Else
                m_intCurrent = 0
            End If

        Catch ex As System.Data.SqlClient.SqlException
            Me.CmnLogOut(Nothing, CST.enmLogType.Error, strSelectString, False)
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :読み込み
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
    Public Function Read() As Boolean
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            'データが取得できていない場合
            If m_intCurrent < 0 Then
                Return False
            End If

            'データが取得できていて、かつ有効行数の範囲内である場合
            'カレントを次の行へ遷移する
            If m_intCurrent < m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows.Count Then
                m_intCurrent = m_intCurrent + 1
                Return True
            End If

            '有効行数の範囲外となった場合
            Return False

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :日付文字列の日付型変換
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
    Public Function CnvSqlDateStringToYMD(ByVal strDate As String) As String
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim strTemp As String

            Select Case strDate.Length
                Case 8
                    '8桁 yyyymmdd → yyyy/mm/dd
                    strTemp = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2)
                Case 14
                    '14桁 yyyymmddhhmmss → yyyy/mm/dd hh:mm:ss
                    strTemp = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                Case 17
                    '17桁 yyyymmddhhmmssmmm → yyyy/mm/dd hh:mm:ss ミリ秒を切り捨て。ミリ秒を格納できないため
                    strTemp = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7, 2) & " " & Mid(strDate, 9, 2) & ":" & Mid(strDate, 11, 2) & ":" & Mid(strDate, 13, 2)
                Case Else
                    '対応していない場合は空文字とする
                    strTemp = ""
            End Select

            Return strTemp

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（String）
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
    Public Function CnvSqlToStr(ByRef strFiled As String _
                     , Optional ByVal strDateFormatString As String = "yyyy/MM/dd") As String
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return ""
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Columns(strFiled).DataType.Name = "DateTime" Then
                '日付タイプの場合
                Return CType(objBuff, DateTime).ToString(strDateFormatString)
            Else
                Return CType(objBuff, String)
            End If

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Date）
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
    Public Function CnvSqlToDte(ByRef strFiled As String) As Date
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return Nothing
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return CType(objBuff, DateTime)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Boolean）
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
    Public Function CnvSqlToBln(ByRef strFiled As String) As Boolean
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return Nothing
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return (CType(objBuff, Integer) <> 0) '0:False 1:True

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Short）
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
    Public Function CnvSqlToSho(ByRef strFiled As String) As Short
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return 0
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return CType(objBuff, Short)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Integer）
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
    Public Function CnvSqlToInt(ByRef strFiled As String) As Integer
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return 0
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Dim strBuff As String
            strBuff = CType(objBuff, String)
            If strBuff.Length = 0 Then
                Return 0
            End If
            Return CType(strBuff, Integer)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Long）
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
    Public Function CnvSqlToLng(ByRef strFiled As String) As Long
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return 0
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return CType(objBuff, Long)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Decimal）
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
    Public Function CnvSqlToDml(ByRef strFiled As String) As Decimal
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return 0
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return CType(objBuff, Decimal)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Double）
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
    Public Function CnvSqlToDbl(ByRef strFiled As String) As Double
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim objBuff As Object

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return 0
            End If
            objBuff = m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled)
            Return CType(objBuff, Double)

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :データ取得（Binary）
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
    Public Function CnvSqlToBny(ByRef strFiled As String) As Byte()
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try
            Dim bteBuff As Byte()

            If m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).IsNull(strFiled) Then
                Return Nothing
            End If
            bteBuff = CType(m_objSqlDataSet.Tables(CST.DataObjName.Table).Rows(m_intCurrent - 1).Item(strFiled), Byte())

            Return bteBuff

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Function

    '**********************************************************************
    ' Name       :クローズ
    '
    '----------------------------------------------------------------------
    ' Comment    :メンバ変数(m_objSqlDataSet)をクローズする。
    '
    '----------------------------------------------------------------------
    ' Design On  :
    '        By  :
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub DBClose()
        Me.CmnLogOut(Nothing, ClsCommon.ClsCst.enmLogType.MethodST, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        Try

            If Not m_objSqlDataSet Is Nothing Then m_objSqlDataSet.Clear()

        Catch ex As Exception
            Me.CmnLogOut(ex, ClsCst.enmLogType.Error, "", False)
            Throw
        Finally
            Me.CmnLogOut(Nothing, CST.enmLogType.MethodED, System.Reflection.MethodBase.GetCurrentMethod().Name, True)
        End Try
    End Sub

    '**********************************************************************
    ' Name       :ログ出力(コモンからのイベントによる)
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
    Public Sub CmnLogOut(ByVal ex As Exception _
                        , ByVal enmLogType As CST.enmLogType _
                        , ByVal strText As String _
                        , ByVal blnDebugLogOutOnly As Boolean)
        If ex Is Nothing Then
            RaiseEvent CLAcsCtrlLogOut(Nothing, enmLogType, strText, blnDebugLogOutOnly)
        Else
            RaiseEvent CLAcsCtrlLogOut(ex, ClsCommon.ClsCst.enmLogType.Error, Nothing, False)
        End If
    End Sub

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
    Public Property SysInf() As Rec_XSys
        Get
            Return m_recXSys
        End Get
        Set(ByVal Value As Rec_XSys)
            m_recXSys = Value
        End Set
    End Property

End Class
