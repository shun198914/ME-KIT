Option Strict On
'**************************************************************************
' Name       :ClsOrderCom 共通ライブラリ
'
'--------------------------------------------------------------------------
' Comment    :
'
'--------------------------------------------------------------------------
' Design On  :FTC)S.Shin
'        By  :2011/11/08
'--------------------------------------------------------------------------
' Update On  :2016/05/06
'        By  :FTC)T.IKEDA
' Comment    :食種によるコメント付加処理で、コメント２つ目と３つめの朝昼夕の区分チェック
'             漏れから、不要なコメントが付加される障害対応
'             朝昼夕うでのIF文チェックを追加した
'**************************************************************************
' Imports
'**************************************************************************
Imports ClsDataFormat
Imports System.Collections.Specialized
'**************************************************************************
' Class
'**************************************************************************
Public Class ClsOrderCom

    Private m_hashWardCdChg As Hashtable         '病棟
    Private m_hashRoomCdChg As Hashtable         '病室
    Private m_hashCoStCdChg As Hashtable         '診療科
    Private m_hashSickCdChg As Hashtable         '病名
    Private m_hashSdFmCdChg As Hashtable         '副食形態
    Private m_hashDlvyCdChg As Hashtable         '配膳先

    Private m_hashMealCdChg As Hashtable         '食種
    Private m_hashWardMealAddCd As Hashtable     '病棟で食種変換用

    Private m_hashMainFdChg As Hashtable         '主食
    Private m_hashMilkFdChg As Hashtable         '調乳食
    Private m_hashTubeFdChg As Hashtable         '経管食

    Private m_hashCmtCdChg As Hashtable          'コメント
    Private m_hashCmtKinCdChg As Hashtable       '禁止コメント

    Private m_hashMealMst As Hashtable           '食種マスタ
    Private m_hashCmtMst As Hashtable            'コメントマスタ

    Private m_hashTubeMealMethod As Hashtable    '経管食の給与方法

    Private m_hashJiCdChg As Hashtable     '事業所コード

    Private Const ORDER_DELIMITER_CHAR_PIPE As String = "|"
    Private Const ORDER_DELIMITER_CHAR_SEMICOLON As String = ";"

    Private Enum enumCmtCls

        NorCmt = 0
        KinCmt = 1

    End Enum

    '**********************************************************************
    ' Name       :初期化
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :FTC)S.Shin
    '        By  :2011/11/08
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Sub New()

        m_hashWardCdChg = New Hashtable
        m_hashRoomCdChg = New Hashtable
        m_hashMealCdChg = New Hashtable
        m_hashMainFdChg = New Hashtable
        m_hashMilkFdChg = New Hashtable
        m_hashTubeFdChg = New Hashtable
        m_hashCmtCdChg = New Hashtable
        m_hashCmtKinCdChg = New Hashtable
        m_hashCoStCdChg = New Hashtable
        m_hashSickCdChg = New Hashtable
        m_hashSdFmCdChg = New Hashtable
        m_hashDlvyCdChg = New Hashtable
        m_hashMealMst = New Hashtable
        m_hashCmtMst = New Hashtable
        m_hashTubeMealMethod = New Hashtable
        m_hashJiCdChg = New Hashtable

        m_hashWardMealAddCd = New Hashtable

    End Sub

    '**********************************************************************
    ' Name       :
    '
    '----------------------------------------------------------------------
    ' Comment    :
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
    Public Function CnvStrToInteger(ByVal strVal As String) As Integer

        Dim intRtnValue As Integer

        Try
            intRtnValue = 0

            Dim dblValue As Double

            If Double.TryParse(strVal, dblValue) Then
                intRtnValue = CType(strVal, Integer)
            Else
                ''数値設定なのに、数値変換できず
                intRtnValue = 0
            End If

            Return intRtnValue

        Catch ex As Exception
            Return -1
        End Try

    End Function

    '**********************************************************************
    ' Name       :
    '
    '----------------------------------------------------------------------
    ' Comment    :
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
    Public Function CnvStrToDouble(ByVal strVal As String) As Double

        Dim dblRtnValue As Double

        Try
            dblRtnValue = 0

            Dim dblValue As Double

            If Double.TryParse(strVal, dblValue) Then
                dblRtnValue = CType(strVal, Double)
            Else
                ''数値設定なのに、数値変換できず
                dblRtnValue = 0
            End If

            Return dblRtnValue

        Catch ex As Exception
            Return -1
        End Try

    End Function


    '**********************************************************************
    ' Name       :
    '
    '----------------------------------------------------------------------
    ' Comment    :
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
    Public Function CnvStrToDecimal(ByVal strVal As String) As Decimal

        Dim dblRtnValue As Decimal

        Try
            dblRtnValue = 0

            Dim dblValue As Double

            If Double.TryParse(strVal, dblValue) Then
                dblRtnValue = CType(strVal, Decimal)
            Else
                ''数値設定なのに、数値変換できず
                dblRtnValue = 0
            End If

            Return dblRtnValue

        Catch ex As Exception
            Return -1
        End Try

    End Function
    '**********************************************************************
    ' Name       :pfOrderDataFmt
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
    Private Function pfOrderDataFmt(ByVal strValue As String, _
                                    ByVal strFmt As String) As String

        Dim strRtnValue As String

        Try
            strRtnValue = strValue

            If strFmt.IndexOf("{NONE}") > -1 Then

                ''フォーマットしない
                strRtnValue = strValue

            ElseIf strFmt.IndexOf("[D]") > -1 Then

                '文字列付加
                strRtnValue = strFmt.Replace("[D]", strValue)

            ElseIf strFmt.IndexOf("{IF}") > -1 Then
                Dim strIfData As String() = strFmt.Replace("{IF}", "").Split("|"c)
                '例 {IF}M:1,F:2 
                'Mであれば1,Fであれば2に変換する
                For Each strIf As String In strIfData
                    If strIf.Split(":"c)(0) = strValue Then
                        strRtnValue = strIf.Split(":"c)(1)
                    End If
                Next

            ElseIf strFmt.IndexOf("{REP}") > -1 Then
                Dim strIfData As String() = strFmt.Replace("{REP}", "").Split("|"c)
                '例 {REP},:、 
                '文字列に,があれば、に変換する
                For Each strIf As String In strIfData
                    strRtnValue = strValue.Replace(strIf.Split(":"c)(0), strIf.Split(":"c)(1))
                Next

            ElseIf strFmt.IndexOf("{ORPC}") > -1 Then
                Dim strIfData As String() = strFmt.Replace("{ORPC}", "").Split("|"c)
                '例 {ORPC},|、 
                '文字列に,があれば、に変換する
                strRtnValue = strValue.Replace(strIfData(0), strIfData(1))

            ElseIf strFmt.IndexOf("{DELCRLF}") > -1 Then

                '例 {DELCRLF} 
                '改行コードがあれば、削除する
                strRtnValue = strValue.Replace(Chr(13), "").Replace(Chr(10), "")

            ElseIf strFmt.IndexOf("{SORT}") > -1 Then

                Dim strSortData As String() = strFmt.Replace("{SORT}", "").Split("|"c)
                '例 {SORT} 

                strRtnValue = ""
                '文字列に番号があればに並び変える
                For Each strSort As String In strSortData

                    If CInt(strSort) < strValue.Length Then
                        strRtnValue &= strValue.Substring(CInt(strSort), 1)
                    End If
                Next

            ElseIf strFmt.IndexOf("s") >= 0 OrElse strFmt.IndexOf("S") >= 0 Then
                '整数の桁数(文字数)
                Dim intIntegerCnt As Integer = strFmt.LastIndexOf("S") + 1
                '小数の桁数(文字数)
                Dim intDecimalCnt As Integer = strFmt.LastIndexOf("s") - intIntegerCnt + 1

                '文字列から小数点位置を取得
                Dim intPoint As Integer = strValue.IndexOf(".")

                Dim strInteger As String
                Dim strDecimal As String


                If intPoint = -1 Then
                    intPoint = strValue.Length
                End If

                strInteger = strValue.Substring(0, intPoint)

                If intPoint < strValue.Length Then
                    strDecimal = strValue.Substring(intPoint + 1, strValue.Length - intPoint - 1)
                Else
                    strDecimal = ""
                End If

                If strInteger.Trim <> "" Then
                    strInteger = Right(strInteger.PadLeft(intIntegerCnt, "0"c), intIntegerCnt)
                End If

                If strDecimal.Trim <> "" Then
                    strDecimal = Left(strDecimal.PadRight(intDecimalCnt, "0"c), intDecimalCnt)
                End If

                strRtnValue = strInteger & strDecimal

            ElseIf strFmt.IndexOf("{LPad}") >= 0 OrElse strFmt.IndexOf("{RPad}") >= 0 Then

                Dim strPadInfo As String()
                Select Case strFmt.Substring(0, 6)
                    Case "{LPad}"
                        strPadInfo = strFmt.Replace("{LPad}", "").Split("|"c)
                        strRtnValue = Trim(strValue).PadLeft(CInt(strPadInfo(1)), CChar(strPadInfo(0)))
                    Case "{RPad}"
                        strPadInfo = strFmt.Replace("{RPad}", "").Split("|"c)
                        strRtnValue = Trim(strValue).PadRight(CInt(strPadInfo(1)), CChar(strPadInfo(0)))
                End Select

            ElseIf strFmt.IndexOf(".") > 0 Then

                '浮動小数点用フォーマット
                Dim intPoint As Integer = strFmt.IndexOf(".")
                strRtnValue = strValue.Substring(0, intPoint) & "." & strValue.Substring(intPoint, Len(strValue) - intPoint)

            Else
                strRtnValue = String.Format(strFmt, strValue)
            End If

            Return strRtnValue

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfOrderDataExceptPoint
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
    Private Function pfOrderDataExceptPoint(ByVal strValue As String, _
                                            ByVal strFmt As String) As String


        Dim strRtnValue As String

        Try
            strRtnValue = strValue

            Dim intValIndex As Integer = -1
            Dim strNum As String = ""           ''整数
            Dim strDec As String = ""           ''小数

            ''整数部を取得
            intValIndex = strFmt.IndexOf(".")

            If strValue <> "" Then
                strNum = CType(Math.Floor(CDbl(strValue)), String)
            Else
                strNum = ""
            End If

            If intValIndex > -1 Then
                strNum = strNum.PadLeft(intValIndex, "0"c)
                strNum = strNum.Substring(0, intValIndex)
            Else
                strNum = strValue
            End If


            ''小数部を取得
            If strValue.IndexOf(".") > -1 Then

                strDec = strValue.ToString.Split("."c)(1)

                If intValIndex > -1 Then

                    intValIndex = strFmt.Length - strFmt.IndexOf(".") - 1

                    strDec = strDec.PadLeft(intValIndex, "0"c)
                    strDec = strDec.Substring(0, intValIndex)
                Else
                    strDec = ""
                End If

            Else

                If intValIndex > -1 Then

                    intValIndex = strFmt.Length - strFmt.IndexOf(".") - 1

                    strDec = strDec.PadLeft(intValIndex, "0"c)
                    strDec = strDec.Substring(0, intValIndex)

                End If

            End If

            strRtnValue = strNum & strDec

            Return strRtnValue

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtSplit
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
    Private Function pfGetCmtSplit(ByVal strValue As String, _
                                   ByVal strDataFmt As String, _
                                   ByVal intChgRule As Integer) As StringCollection

        Dim strRtnValue As New StringCollection

        Try

            Dim objStrWkStringCol As New StringCollection
            Dim strSplitData As String() = strDataFmt.Replace("{NONE}", "").Split("|"c)
            Dim chrSplit As Char

            objStrWkStringCol.Add(strValue)

            For intLoopCnt As Integer = 0 To strSplitData.Count - 1

                chrSplit = CType(strSplitData(intLoopCnt), Char)

                Dim objStrWkStringCol2 As New StringCollection

                For Each strWkStr As String In objStrWkStringCol

                    Dim strDataCol As String() = strWkStr.Split(chrSplit)

                    For Each strData As String In strDataCol

                        objStrWkStringCol2.Add(strData)

                    Next

                Next

                objStrWkStringCol.Clear()

                For Each strWkStr As String In objStrWkStringCol2

                    objStrWkStringCol.Add(strWkStr)

                Next

            Next

            If intChgRule = 0 Then

                For Each strWkStr As String In objStrWkStringCol

                    strRtnValue.Add(strWkStr)

                Next

            Else


                For intLoopCnt As Integer = intChgRule - 1 To objStrWkStringCol.Count - 1 Step strSplitData.Count

                    strRtnValue.Add(objStrWkStringCol.Item(intLoopCnt))

                Next

            End If

            Return strRtnValue

        Catch ex As Exception

            Return New StringCollection

        End Try

    End Function

    '**********************************************************************
    ' Name       :OrderDataFmt
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
    Public Function OrderDataFmt(ByVal strValue As String, _
                                 ByVal strFmt As String) As String

        Try
            Return pfOrderDataFmt(strValue, strFmt)

        Catch ex As Exception
            Return ""
        End Try

    End Function

    '**********************************************************************
    ' Name       :OrderDataFmt
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
    Public Function OrderDataExceptPoint(ByVal strValue As String, _
                                         ByVal strFmt As String) As String

        Try
            Return pfOrderDataExceptPoint(strValue, strFmt)

        Catch ex As Exception
            Return ""
        End Try

    End Function


    '**********************************************************************
    ' Name       :オーダ連携共通データ変換処理
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
    Public Function OrderDataConvertSTD(ByVal objRecORecv As Rec_ORecv, _
                                        ByVal objColOTelDef As Col_OTelDef, _
                                        ByVal SysInf As Rec_XSys) As Rec_OrderSTD

        Dim objOrderSTD As New Rec_OrderSTD
        Dim strAjiName As String = ""

        Try

            With objOrderSTD

                .OpYMD = objRecORecv.OpYMD                     '登録年月日()
                .OpNo = objRecORecv.OpNo                       '登録番号()
                .OdTp = objRecORecv.OdTp                       'オーダ種別()
                .RfctDiv = objRecORecv.RfctDiv                 '反映区分()
                .PpIsDiv = objRecORecv.PpIsDiv                 '食事箋発行区分()
                .OdNo = objRecORecv.OdNo                       'オーダ番号()
                .Rsheet = objRecORecv.Rsheet                   '食事箋/伝票コード
                .AtvOd = objRecORecv.AtvOd                     '実施オーダ区分()
                .Deadline = objRecORecv.Deadline               '締切り区分()
                .ExecDiv = objRecORecv.ExecDiv                 '処理区分()
                .TelMDiv = objRecORecv.TelMDiv                 '電文種別()
                .OdDiv = objRecORecv.OdDiv                     'オーダ指示区分()
                .UpDtUser = objRecORecv.UpDtUser               '更新者

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                'オーダデータの読込み
                'エンコードオブジェクトを用意します。
                Dim strEncoding As String
                Dim objEnc As Text.Encoding

                'エンコードを取得
                strEncoding = SysInf.OrderRcvEncoding

                If strEncoding = "" Then
                    strEncoding = "Shift_Jis"   ''shift_jis
                End If

                '文字列をstrEncodingで変換した文字にします。
                objEnc = Text.Encoding.GetEncoding(strEncoding)

                ''オーダデータを取得
                Dim bytOrderData As Byte()
                bytOrderData = objEnc.GetBytes(objRecORecv.OdData)

                Dim strValue As String = ""

                Dim intFlexStart As Integer = 0
                Dim intFlexLoop As Integer = 0

                Dim objFlexLoop As New Hashtable


                ''電文定義をループ
                For Each objRecOTelDef As Rec_OTelDef In objColOTelDef

                    strAjiName = objRecOTelDef.AjiNm

                    '開始位置
                    Dim intStart As Integer = objRecOTelDef.ByteSt

                    '桁数
                    Dim intDigit As Integer = objRecOTelDef.DataLen

                    intFlexLoop = 1

                    ''可変長の場合
                    If objRecOTelDef.Flexible = 1 Then

                        If objRecOTelDef.ByteSt = -1 Then
                            intStart = intFlexStart
                        End If

                        ''ループする値がある場合
                        If objFlexLoop.ContainsKey(objRecOTelDef.AjiNm) = True Then
                            intFlexLoop = CType(objFlexLoop.Item(objRecOTelDef.AjiNm).ToString, Integer)
                        End If

                    Else
                        intFlexStart = intStart
                    End If

                    ''可変長対応
                    For intLoopCnt As Integer = 0 To intFlexLoop - 1

                        ''開始と長さが共に０の場合は規定値を使用()
                        If intStart = 0 AndAlso intDigit = 0 Then

                            '規定値を格納する
                            strValue = objRecOTelDef.DefVal

                        Else

                            intStart = intFlexStart

                            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                            'オーダのデータから必要な部分を取り出す
                            If bytOrderData.Length >= intStart + intDigit Then
                                strValue = objEnc.GetString(bytOrderData, intStart, intDigit)
                            Else

                                .ErrMsg = "電文バイト数より大きいサイズの文字変換が発生しました。Column=" & objRecOTelDef.AjiNm & _
                                                                                               " Start=" & objRecOTelDef.ByteSt & _
                                                                                               " DIGIT=" & objRecOTelDef.DataLen
                                Return objOrderSTD

                            End If
                            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                        End If

                        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                        ''値がない場合は規定値
                        If strValue.Trim = "" Then

                            '規定値を格納する
                            strValue = objRecOTelDef.DefVal

                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                        'オーダのデータが数値がどうかを確認する
                        Dim blnNumericError As Boolean

                        If objRecOTelDef.DataTp = "Integer" OrElse _
                           objRecOTelDef.DataTp = "Long" OrElse _
                           objRecOTelDef.DataTp = "Double" Then

                            Dim dblValue As Double

                            If Double.TryParse(strValue, dblValue) Then
                                blnNumericError = False
                            Else
                                ''数値設定なのに、数値変換できず
                                blnNumericError = True
                            End If
                        Else
                            blnNumericError = False
                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／


                        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                        'フォーマット
                        If objRecOTelDef.DataFmt <> "" Then

                            ''空白の場合はフォーマットしない
                            If strValue.Trim <> "" Then
                                strValue = pfOrderDataFmt(strValue, objRecOTelDef.DataFmt)
                            End If

                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／


                        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                        '値の格納
                        If strValue.Trim <> "" AndAlso blnNumericError = False Then

                            ''値が空白じゃなかった場合
                            Select Case objRecOTelDef.AjiNm

                                Case "OpYMD"
                                    .OpYMD = strValue
                                Case "OpNo"
                                    .OpNo = CType(strValue, Integer)
                                Case "SeqNo"
                                    .SeqNo = CType(strValue, Integer)
                                Case "OpNoRef"
                                    .OpNoRef = strValue
                                Case "OdTp"
                                    .OdTp = strValue
                                Case "RfctDiv"
                                    .RfctDiv = strValue
                                Case "PpIsDiv"
                                    .PpIsDiv = CType(strValue, Integer)
                                Case "OdNo"
                                    .OdNo = strValue
                                Case "Rsheet"
                                    .Rsheet = strValue
                                Case "AtvOd"
                                    .AtvOd = strValue
                                Case "Deadline"
                                    .Deadline = strValue
                                Case "ExecDiv"
                                    .ExecDiv = strValue
                                Case "TelMDiv"
                                    .TelMDiv = strValue
                                Case "OdDiv"
                                    .OdDiv = strValue
                                Case "ComOdDiv"
                                    .ComOdDiv = strValue

                                Case "IncrementNo"
                                    .IncrementNo = strValue
                                Case "SystmCd"
                                    .SystmCd = strValue
                                Case "RecFlg"
                                    .RecFlg = strValue
                                Case "RecvPcId"
                                    .ReceivePcName = strValue
                                Case "SenderPcId"
                                    .SenderPcName = strValue
                                Case "InputYMD"
                                    .InputYMD = strValue
                                Case "InputTime"
                                    .InputTime = strValue
                                Case "Terminal"
                                    .Terminal = strValue
                                Case "UserNo"
                                    .UserNo = strValue
                                Case "AnsType"
                                    .AnsType = strValue
                                Case "DataLeg"
                                    .DataLength = strValue

                                Case "JiCd"
                                    .JiCd = pfChangJiCd(strValue)

                                Case "PerCd"
                                    .PerCd = strValue
                                Case "StYMD"
                                    .StYMD = strValue
                                Case "StTmDiv"
                                    .StTmDiv = CType(strValue, Integer)
                                Case "EdYMD"
                                    .EdYMD = strValue
                                Case "EdTmDiv"
                                    .EdTmDiv = CType(strValue, Integer)
                                Case "MovYMD"
                                    .MovYMD = strValue
                                Case "MovTime"
                                    .MovTime = strValue
                                Case "PerNm"
                                    .PerNm = strValue
                                Case "PerKana"
                                    .PerKana = strValue
                                Case "Sex"
                                    .Sex = CType(strValue, Integer)
                                Case "Birth"
                                    .Birth = strValue
                                Case "RefColM"
                                    .RefColM = strValue
                                Case "PerHigh"
                                    .PerHigh = CType(strValue, Double)
                                Case "PerWeit"
                                    .PerWeit = CType(strValue, Double)
                                Case "LifeLv"
                                    .LifeLv = strValue
                                Case "MedCare"
                                    .MedCare = strValue
                                Case "PDoctor"
                                    .PDoctor = strValue
                                Case "SDoctor"
                                    .SDoctor = strValue
                                Case "NstDiv"
                                    .NstDiv = CType(strValue, Integer)
                                Case "RefColFM"
                                    .RefColFM = strValue
                                Case "RefColF"
                                    .RefColF = strValue
                                Case "RefColF10"
                                    .RefColF10 = strValue
                                Case "RefColFL"
                                    .RefColFL = strValue
                                Case "RefColF15"
                                    .RefColF15 = strValue
                                Case "RefColFE"
                                    .RefColFE = strValue
                                Case "RefColF20"
                                    .RefColF20 = strValue
                                Case "Ward"
                                    .Ward = strValue
                                Case "WardM"
                                    .WardM = strValue
                                Case "Ward10"
                                    .Ward10 = strValue
                                Case "WardL"
                                    .WardL = strValue
                                Case "Ward15"
                                    .Ward15 = strValue
                                Case "WardE"
                                    .WardE = strValue
                                Case "Ward20"
                                    .Ward20 = strValue
                                Case "Room"
                                    .Room = strValue
                                Case "RoomM"
                                    .RoomM = strValue
                                Case "Room10"
                                    .Room10 = strValue
                                Case "RoomL"
                                    .RoomL = strValue
                                Case "Room15"
                                    .Room15 = strValue
                                Case "RoomE"
                                    .RoomE = strValue
                                Case "Room20"
                                    .Room20 = strValue
                                Case "Bed"
                                    .Bed = strValue
                                Case "BedM"
                                    .BedM = strValue
                                Case "Bed10"
                                    .Bed10 = strValue
                                Case "BedL"
                                    .BedL = strValue
                                Case "Bed15"
                                    .Bed15 = strValue
                                Case "BedE"
                                    .BedE = strValue
                                Case "Bed20"
                                    .Bed20 = strValue
                                Case "NoMeal"
                                    .NoMeal = CType(strValue, Integer)
                                Case "NoMealM"
                                    .NoMealM = CType(strValue, Integer)
                                Case "NoMeal10"
                                    .NoMeal10 = CType(strValue, Integer)
                                Case "NoMealL"
                                    .NoMealL = CType(strValue, Integer)
                                Case "NoMeal15"
                                    .NoMeal15 = CType(strValue, Integer)
                                Case "NoMealE"
                                    .NoMealE = CType(strValue, Integer)
                                Case "NoMeal20"
                                    .NoMeal20 = CType(strValue, Integer)
                                Case "NoMlRs"
                                    .NoMlRs = strValue
                                Case "NoMlRsM"
                                    .NoMlRsM = strValue
                                Case "NoMlRs10"
                                    .NoMlRs10 = strValue
                                Case "NoMlRsL"
                                    .NoMlRsL = strValue
                                Case "NoMlRs15"
                                    .NoMlRs15 = strValue
                                Case "NoMlRsE"
                                    .NoMlRsE = strValue
                                Case "NoMlRs20"
                                    .NoMlRs20 = strValue
                                Case "MealCd"
                                    .MealCd = strValue
                                Case "MealCdM"
                                    .MealCdM = strValue
                                Case "MealCd10"
                                    .MealCd10 = strValue
                                Case "MealCdL"
                                    .MealCdL = strValue
                                Case "MealCd15"
                                    .MealCd15 = strValue
                                Case "MealCdE"
                                    .MealCdE = strValue
                                Case "MealCd20"
                                    .MealCd20 = strValue
                                Case "MaCkCd"
                                    .MaCkCd = strValue
                                Case "MaCkCdM"
                                    .MaCkCdM = strValue
                                Case "MaCkCd10"
                                    .MaCkCd10 = strValue
                                Case "MaCkCdL"
                                    .MaCkCdL = strValue
                                Case "MaCkCd15"
                                    .MaCkCd15 = strValue
                                Case "MaCkCdE"
                                    .MaCkCdE = strValue
                                Case "MaCkCd20"
                                    .MaCkCd20 = strValue
                                Case "SdShap"
                                    .SdShap = strValue
                                Case "SdShapM"
                                    .SdShapM = strValue
                                Case "SdShap10"
                                    .SdShap10 = strValue
                                Case "SdShapL"
                                    .SdShapL = strValue
                                Case "SdShap15"
                                    .SdShap15 = strValue
                                Case "SdShapE"
                                    .SdShapE = strValue
                                Case "SdShap20"
                                    .SdShap20 = strValue
                                Case "EatAdDiv"
                                    .EatAdDiv = CType(strValue, Integer)
                                Case "EatAdDivM"
                                    .EatAdDivM = CType(strValue, Integer)
                                Case "EatAdDiv10"
                                    .EatAdDiv10 = CType(strValue, Integer)
                                Case "EatAdDivL"
                                    .EatAdDivL = CType(strValue, Integer)
                                Case "EatAdDiv15"
                                    .EatAdDiv15 = CType(strValue, Integer)
                                Case "EatAdDivE"
                                    .EatAdDivE = CType(strValue, Integer)
                                Case "EatAdDiv20"
                                    .EatAdDiv20 = CType(strValue, Integer)
                                Case "MCPrt"
                                    .MCPrt = CType(strValue, Integer)
                                Case "MCPrtM"
                                    .MCPrtM = CType(strValue, Integer)
                                Case "MCPrt10"
                                    .MCPrt10 = CType(strValue, Integer)
                                Case "MCPrtL"
                                    .MCPrtL = CType(strValue, Integer)
                                Case "MCPrt15"
                                    .MCPrt15 = CType(strValue, Integer)
                                Case "MCPrtE"
                                    .MCPrtE = CType(strValue, Integer)
                                Case "MCPrt20"
                                    .MCPrt20 = CType(strValue, Integer)
                                Case "SelEtDivDef"
                                    .SelEtDivDef = CType(strValue, Integer)
                                Case "SelEtDivDefM"
                                    .SelEtDivDefM = CType(strValue, Integer)
                                Case "SelEtDivDef10"
                                    .SelEtDivDef10 = CType(strValue, Integer)
                                Case "SelEtDivDefL"
                                    .SelEtDivDefL = CType(strValue, Integer)
                                Case "SelEtDivDef15"
                                    .SelEtDivDef15 = CType(strValue, Integer)
                                Case "SelEtDivDefE"
                                    .SelEtDivDefE = CType(strValue, Integer)
                                Case "SelEtDivDef20"
                                    .SelEtDivDef20 = CType(strValue, Integer)
                                Case "DelivCd"
                                    .DelivCd = strValue
                                Case "DelivCdM"
                                    .DelivCdM = strValue
                                Case "DelivCd10"
                                    .DelivCd10 = strValue
                                Case "DelivCdL"
                                    .DelivCdL = strValue
                                Case "DelivCd15"
                                    .DelivCd15 = strValue
                                Case "DelivCdE"
                                    .DelivCdE = strValue
                                Case "DelivCd20"
                                    .DelivCd20 = strValue
                                Case "SelEtDiv"
                                    .SelEtDiv = strValue

                                Case "CmtCdCOM"
                                    .CmtCd.Add(strValue)
                                Case "CmtCdM"
                                    .CmtCdM.Add(strValue)
                                Case "CmtCd10"
                                    .CmtCd10.Add(strValue)
                                Case "CmtCdL"
                                    .CmtCdL.Add(strValue)
                                Case "CmtCd15"
                                    .CmtCd15.Add(strValue)
                                Case "CmtCdE"
                                    .CmtCdE.Add(strValue)
                                Case "CmtCd20"
                                    .CmtCd20.Add(strValue)

                                Case "CmtKinCdCOM"
                                    .CmtKinCd.Add(strValue)
                                Case "CmtKinCdM"
                                    .CmtKinCdM.Add(strValue)
                                Case "CmtKinCd10"
                                    .CmtKinCd10.Add(strValue)
                                Case "CmtKinCdL"
                                    .CmtKinCdL.Add(strValue)
                                Case "CmtKinCd15"
                                    .CmtKinCd15.Add(strValue)
                                Case "CmtKinCdE"
                                    .CmtKinCdE.Add(strValue)
                                Case "CmtKinCd20"
                                    .CmtKinCd20.Add(strValue)

                                Case "CmtSplitCOM"

                                    Dim objWkValue As New StringCollection

                                    objWkValue = pfGetCmtSplit(strValue, objRecOTelDef.DataFmt, objRecOTelDef.ChgRule)

                                    For Each strWkValue As String In objWkValue
                                        .CmtCd.Add(strWkValue)
                                    Next

                                Case "CmtKinSplitCOM"

                                    Dim objWkValue As New StringCollection

                                    objWkValue = pfGetCmtSplit(strValue, objRecOTelDef.DataFmt, objRecOTelDef.ChgRule)

                                    For Each strWkValue As String In objWkValue
                                        .CmtKinCd.Add(strValue)
                                    Next

                                Case "NtCd"
                                    .NtCd.Add(objRecOTelDef.OrdColNm)
                                    .LmtVal.Add(strValue)

                                Case "NtCdM"
                                    .NtCdM.Add(objRecOTelDef.OrdColNm)
                                    .LmtValM.Add(strValue)

                                Case "NtCdL"
                                    .NtCdL.Add(objRecOTelDef.OrdColNm)
                                    .LmtValL.Add(strValue)

                                Case "NtCdE"
                                    .NtCdE.Add(objRecOTelDef.OrdColNm)
                                    .LmtValE.Add(strValue)

                                Case "NutCd"
                                    .NutCd.Add(strValue)
                                Case "NutMealCnt"
                                    .NutMealCnt.Add(strValue)
                                Case "NutMeal"
                                    .NutMeal.Add(strValue)
                                Case "NutOneVal"
                                    .NutOneVal.Add(strValue)
                                Case "NutDlvVal"
                                    .NutDlvVal.Add(strValue)
                                Case "NutUseVal"
                                    .NutUseVal.Add(strValue)
                                Case "NutRefCol"
                                    .NutRefCol.Add(strValue)

                                Case "NutCdM"
                                    .NutCdM.Add(strValue)
                                Case "NutMealCntM"
                                    .NutMealCntM.Add(strValue)
                                Case "NutMealM"
                                    .NutMealM.Add(strValue)
                                Case "NutOneValM"
                                    .NutOneValM.Add(strValue)
                                Case "NutDlvValM"
                                    .NutDlvValM.Add(strValue)
                                Case "NutUseValM"
                                    .NutUseValM.Add(strValue)
                                Case "NutRefColM"
                                    .NutRefColM.Add(strValue)

                                Case "NutCdL"
                                    .NutCdL.Add(strValue)
                                Case "NutMealCntL"
                                    .NutMealCntL.Add(strValue)
                                Case "NutMealL"
                                    .NutMealL.Add(strValue)
                                Case "NutOneValL"
                                    .NutOneValL.Add(strValue)
                                Case "NutDlvValL"
                                    .NutDlvValL.Add(strValue)
                                Case "NutUseValL"
                                    .NutUseValL.Add(strValue)
                                Case "NutRefColL"
                                    .NutRefColL.Add(strValue)

                                Case "NutCdE"
                                    .NutCdE.Add(strValue)
                                Case "NutMealCntE"
                                    .NutMealCntE.Add(strValue)
                                Case "NutMealE"
                                    .NutMealE.Add(strValue)
                                Case "NutOneValE"
                                    .NutOneValE.Add(strValue)
                                Case "NutDlvValE"
                                    .NutDlvValE.Add(strValue)
                                Case "NutUseValE"
                                    .NutUseValE.Add(strValue)
                                Case "NutRefColE"
                                    .NutRefColE.Add(strValue)

                                Case "NutTimeTm"
                                    .NutTimeTm.Add(strValue)
                                Case "NutCdTime"
                                    .NutCdTime.Add(strValue)
                                Case "NutMealCntTime"
                                    .NutMealCntTime.Add(strValue)
                                Case "NutMealTime"
                                    .NutMealTime.Add(strValue)
                                Case "NutOneValTime"
                                    .NutOneValTime.Add(strValue)
                                Case "NutDlvValTime"
                                    .NutDlvValTime.Add(strValue)
                                Case "NutUseValTime"
                                    .NutUseValTime.Add(strValue)
                                Case "NutRefColTime"
                                    .NutRefColTime.Add(strValue)

                                Case "MilkCd"
                                    .MilkCd.Add(strValue)
                                Case "MilkCon"
                                    .MilkCon.Add(strValue)
                                Case "MilkVal"
                                    .MilkVal.Add(strValue)
                                Case "MilkNum"
                                    .MilkNum.Add(strValue)
                                Case "MilkMealCnt"
                                    .MilkMealCnt.Add(strValue)
                                Case "BabyBot"
                                    .BabyBot.Add(strValue)
                                Case "BotNum"
                                    .BotNum.Add(strValue)
                                Case "Nipple"
                                    .Nipple.Add(strValue)
                                Case "NippNum"
                                    .NippNum.Add(strValue)
                                Case "Dilution"
                                    .Dilution.Add(strValue)
                                Case "MilkRefCol"
                                    .MilkRefCol.Add(strValue)


                                Case "MilkCdM"
                                    .MilkCdM.Add(strValue)
                                Case "MilkConM"
                                    .MilkConM.Add(strValue)
                                Case "MilkValM"
                                    .MilkValM.Add(strValue)
                                Case "MilkNumM"
                                    .MilkNumM.Add(strValue)
                                Case "MilkMealCntM"
                                    .MilkMealCntM.Add(strValue)
                                Case "BabyBotM"
                                    .BabyBotM.Add(strValue)
                                Case "BotNumM"
                                    .BotNumM.Add(strValue)
                                Case "NippleM"
                                    .NippleM.Add(strValue)
                                Case "NippNumM"
                                    .NippNumM.Add(strValue)
                                Case "DilutionM"
                                    .DilutionM.Add(strValue)
                                Case "MilkRefColM"
                                    .MilkRefColM.Add(strValue)

                                Case "MilkCdL"
                                    .MilkCdL.Add(strValue)
                                Case "MilkConL"
                                    .MilkConL.Add(strValue)
                                Case "MilkValL"
                                    .MilkValL.Add(strValue)
                                Case "MilkNumL"
                                    .MilkNumL.Add(strValue)
                                Case "MilkMealCntL"
                                    .MilkMealCntL.Add(strValue)
                                Case "BabyBotL"
                                    .BabyBotL.Add(strValue)
                                Case "BotNumL"
                                    .BotNumL.Add(strValue)
                                Case "NippleL"
                                    .NippleL.Add(strValue)
                                Case "NippNumL"
                                    .NippNumL.Add(strValue)
                                Case "DilutionL"
                                    .DilutionL.Add(strValue)
                                Case "MilkRefColL"
                                    .MilkRefColL.Add(strValue)

                                Case "MilkCdE"
                                    .MilkCdE.Add(strValue)
                                Case "MilkConE"
                                    .MilkConE.Add(strValue)
                                Case "MilkValE"
                                    .MilkValE.Add(strValue)
                                Case "MilkNumE"
                                    .MilkNumE.Add(strValue)
                                Case "MilkMealCntE"
                                    .MilkMealCntE.Add(strValue)
                                Case "BabyBotE"
                                    .BabyBotE.Add(strValue)
                                Case "BotNumE"
                                    .BotNumE.Add(strValue)
                                Case "NippleE"
                                    .NippleE.Add(strValue)
                                Case "NippNumE"
                                    .NippNumE.Add(strValue)
                                Case "DilutionE"
                                    .DilutionE.Add(strValue)
                                Case "MilkRefColE"
                                    .MilkRefColE.Add(strValue)

                                Case "VenaCd"
                                    .VenaCd.Add(strValue)
                                Case "VenaMealCnt"
                                    .VenaMealCnt.Add(strValue)
                                Case "VenaMeal"
                                    .VenaMeal.Add(strValue)
                                Case "VenaOneVal"
                                    .VenaOneVal.Add(strValue)
                                Case "VenaDlvVal"
                                    .VenaDlvVal.Add(strValue)
                                Case "VenaUseVal"
                                    .VenaUseVal.Add(strValue)
                                Case "VenaRefCol"
                                    .VenaRefCol.Add(strValue)

                                Case "IllCd"
                                    .IllCd.Add(strValue)
                                Case "IllCdM"
                                    .IllCdM.Add(strValue)
                                Case "IllCd10"
                                    .IllCd10.Add(strValue)
                                Case "IllCdL"
                                    .IllCdL.Add(strValue)
                                Case "IllCd15"
                                    .IllCd15.Add(strValue)
                                Case "IllCdE"
                                    .IllCdE.Add(strValue)
                                Case "IllCd20"
                                    .IllCd20.Add(strValue)

                                Case "Snack"
                                    .Snack = CType(strValue, Integer)
                                Case "Snack10"
                                    .Snack10 = CType(strValue, Integer)
                                Case "Snack15"
                                    .Snack15 = CType(strValue, Integer)
                                Case "Snack20"
                                    .Snack20 = CType(strValue, Integer)

                                Case "ValidEatTmDivM"
                                    .ValidEatTmDivM = CType(strValue, Integer)
                                Case "ValidEatTmDivL"
                                    .ValidEatTmDivL = CType(strValue, Integer)
                                Case "ValidEatTmDivE"
                                    .ValidEatTmDivE = CType(strValue, Integer)

                                Case "KCnt"
                                    .KCnt = CType(strValue, Integer)
                                Case "KCntM"
                                    .KCntM = CType(strValue, Integer)
                                Case "KCnt10"
                                    .KCnt10 = CType(strValue, Integer)
                                Case "KCntL"
                                    .KCntL = CType(strValue, Integer)
                                Case "KCnt15"
                                    .KCnt15 = CType(strValue, Integer)
                                Case "KCntE"
                                    .KCntE = CType(strValue, Integer)
                                Case "KCnt20"
                                    .KCnt20 = CType(strValue, Integer)
                                Case "KCntWeek"
                                    .KCntWeek = CType(strValue, String)
                                Case "HaveMilkTube"
                                    .HaveMilkTube = CType(strValue, Integer)

                            End Select

                            ''変換ルールを追加する
                            Select Case objRecOTelDef.ChgRule

                                Case 1
                                    objOrderSTD.ChgRule1 &= strValue.Trim
                                Case 2
                                    objOrderSTD.ChgRule2 &= strValue.Trim
                            End Select

                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                        ''可変長ループ項目がある場合
                        If objRecOTelDef.FlexibleLoopTgt <> "" Then

                            If strValue <> "" Then
                                objFlexLoop.Add(objRecOTelDef.FlexibleLoopTgt, strValue)
                            End If

                        End If

                        intFlexStart = intStart + intDigit

                    Next

                Next
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End With


            Return objOrderSTD

        Catch ex As Exception

            objOrderSTD.ErrMsg = "Column=" & strAjiName & " MSG=" & ex.Message
            Return objOrderSTD

        End Try

    End Function

    '**********************************************************************
    ' Summary    : 年齢文字列作成
    '----------------------------------------------------------------------
    ' Method     : Public Function GetPerAgeStr
    '
    ' Parameter  : ByVal strBirth As String....誕生日文字列
    '
    ' Return     :  
    '
    '----------------------------------------------------------------------
    ' Comment    : 引数で渡された誕生日としての日付文字列を元に経過年又はヶ月を付加した年齢文字列を返します。
    '              strBaseDateに日付文字列が指定されている場合は、その日付時点での年齢文字列を返します。
    '              誕生日文字列がNothingや不正な日付、システム日付に対して未来日の場合は「？」を、
    '              ０歳未満の場合は「ｎヶ月」を、
    '              １歳以上の場合は「ｎ歳」を返します。
    '----------------------------------------------------------------------
    ' Design On  : 2008/03/03
    '        By  : FMS)H.Sakamoto
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function GetPerAgeInt(ByVal strBirth As String _
                                , Optional ByVal strBaseDate As String = Nothing) As Integer
        Try
            '--------------------
            '年齢を経過月値で取得します。
            Dim lngPersonalAge As Long = ClsCom.GetPerAge(strBirth, DateInterval.Month, strBaseDate)

            Select Case lngPersonalAge
                Case -1
                    Return -1  '誕生日が不正な日付かエラー
                Case Is < 12
                    Return 0 '０歳児
                Case Else
                    Return CType(Math.Floor(lngPersonalAge / 12), Integer) '１歳以上
            End Select
            '--------------------
        Catch ex As Exception
            Return -1
        Finally

        End Try
    End Function

    '**********************************************************************
    ' Summary    : コメント組合せ関数
    '----------------------------------------------------------------------
    ' Method     : Public Function CommentCombinationCollection
    '
    ' Parameter  : 
    '
    ' Return     :  
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '----------------------------------------------------------------------
    ' Design On  : 2013/09/02
    '        By  : FTC)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function CommentCombinationCollection(ByVal intCmtMax As Integer, _
                                                  ByVal intCmtCount As Integer, _
                                                  ByVal objCmtCol As StringCollection, _
                                         Optional ByVal intCmtMin As Integer = 1, _
                                         Optional ByVal strCmtPreceding As String = "") As StringCollection
        Try

            Dim intCntI As Integer
            Dim Buffer As String
            Dim objBuffer As New StringCollection

            If intCmtMax - intCmtMin + 1 < intCmtCount Then

                ' 「4～5 の中の 7個の組み合わせ」 みたいにアリエナイ場合の処理
                CommentCombinationCollection = Nothing

            ElseIf intCmtCount = 0 Then

                '「0 個の組み合わせ」 の場合
                objBuffer.Add(strCmtPreceding)

                CommentCombinationCollection = objBuffer

            ElseIf intCmtMax - intCmtMin + 1 = intCmtCount Then
                ' 「4～5 の中の 2個の組み合わせ」 みたいな場合の処理
                ' 残りの数字の個数が 指定の個数 と同じだった場合

                Buffer = strCmtPreceding
                For intCntI = intCmtMin To intCmtMax
                    Buffer = Buffer & objCmtCol(intCntI - 1).PadRight(8)
                Next intCntI

                objBuffer.Add(Buffer)

                CommentCombinationCollection = objBuffer

            Else

                objBuffer.Clear()

                Dim objWk As StringCollection

                For intCntI = intCmtMin To intCmtMax - (intCmtCount - 1)
                    objWk = CommentCombinationCollection(intCmtMax, intCmtCount - 1, objCmtCol, intCntI + 1, strCmtPreceding & objCmtCol(intCntI - 1).PadRight(8))
                    For Each strwk As String In objWk
                        objBuffer.Add(strwk)
                    Next
                Next intCntI

                CommentCombinationCollection = objBuffer

            End If

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    '**********************************************************************
    ' Name       :オーダされたコメントをソートする
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
    Private Function pfSortOrderCmt(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                    ByVal blnDelimiterKin As Boolean, _
                                    ByVal objCmtKinMHash As Hashtable, _
                                    ByVal objCmtMHash As Hashtable, _
                                    ByVal strEatTmDiv As String) As StringCollection

        Try

            Dim objOrderCmt As New StringCollection
            Dim objSortOrderCmt As New SortedList

            Dim objRecMCmt As Rec_MCmt

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼

                ''禁止コメント
                If Not objRecOrderSTD.CmtKinCd Is Nothing AndAlso objRecOrderSTD.CmtKinCd.Count > 0 Then

                    For Each strWkCmt As String In objRecOrderSTD.CmtKinCd

                        strWkCmt = strWkCmt.Trim

                        If objCmtKinMHash.ContainsKey(strWkCmt) Then

                            objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                            Select Case strEatTmDiv

                                Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.AM10Div, String)

                                    If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.PM15Div, String)

                                    If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.NightDiv, String)

                                    If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case Else

                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If

                            End Select

                        End If

                    Next

                End If

                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else


                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメント一緒に送ってくる場合 　　　　　　　　　　　＼

                ''禁止コメント
                If Not objRecOrderSTD.CmtKinCd Is Nothing AndAlso objRecOrderSTD.CmtKinCd.Count > 0 Then

                    For Each strWkCmt As String In objRecOrderSTD.CmtKinCd

                        strWkCmt = strWkCmt.Trim

                        If objCmtKinMHash.ContainsKey(strWkCmt) Then

                            objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)


                            Select Case strEatTmDiv

                                Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.AM10Div, String)

                                    If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.PM15Div, String)

                                    If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.NightDiv, String)

                                    If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case Else

                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If

                            End Select

                        End If

                    Next

                End If

                ''通常コメント
                If Not objRecOrderSTD.CmtCd Is Nothing AndAlso objRecOrderSTD.CmtCd.Count > 0 Then

                    For Each strWkCmt As String In objRecOrderSTD.CmtCd

                        ''禁止コメントも見る
                        strWkCmt = strWkCmt.Trim

                        If objCmtKinMHash.ContainsKey(strWkCmt) Then

                            objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                            Select Case strEatTmDiv

                                Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.AM10Div, String)

                                    If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.PM15Div, String)

                                    If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case CType(ClsCst.enmEatTimeDiv.NightDiv, String)

                                    If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Case Else

                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If

                            End Select

                        Else
                            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                            '／'禁止コメントがない場合　　　　　　 　　　　　　　　　　　＼
                            ''通常コメントもみる
                            strWkCmt = strWkCmt.Trim

                            If objCmtMHash.ContainsKey(strWkCmt) Then

                                objRecMCmt = CType(objCmtMHash.Item(strWkCmt), Rec_MCmt)

                                Select Case strEatTmDiv

                                    Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                        If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.AM10Div, String)

                                        If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.PM15Div, String)

                                        If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.NightDiv, String)

                                        If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    Case Else

                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If

                                End Select

                            End If

                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            Select Case strEatTmDiv

                Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                    ''朝
                    If blnDelimiterKin = False Then

                        ''通常コメント
                        If Not objRecOrderSTD.CmtCdM Is Nothing AndAlso objRecOrderSTD.CmtCdM.Count > 0 Then

                            For Each strWkCmt As String In objRecOrderSTD.CmtCdM

                                strWkCmt = strWkCmt.Trim

                                If objCmtKinMHash.ContainsKey(strWkCmt) Then

                                    objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If
                                Else

                                    If objCmtMHash.ContainsKey(strWkCmt) Then

                                        objRecMCmt = CType(objCmtMHash.Item(strWkCmt), Rec_MCmt)

                                        If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    End If
                                End If

                            Next

                        End If

                    End If

                    ''禁止コメント
                    If Not objRecOrderSTD.CmtKinCdM Is Nothing AndAlso objRecOrderSTD.CmtKinCdM.Count > 0 Then

                        For Each strWkCmt As String In objRecOrderSTD.CmtKinCdM

                            strWkCmt = strWkCmt.Trim

                            If objCmtKinMHash.ContainsKey(strWkCmt) Then

                                objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If
                                End If

                            End If

                        Next

                    End If


                Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                    ''昼
                    If blnDelimiterKin = False Then

                        If Not objRecOrderSTD.CmtCdL Is Nothing AndAlso objRecOrderSTD.CmtCdL.Count > 0 Then

                            For Each strWkCmt As String In objRecOrderSTD.CmtCdL

                                strWkCmt = strWkCmt.Trim

                                If objCmtKinMHash.ContainsKey(strWkCmt) Then
                                    objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Else

                                    If objCmtMHash.ContainsKey(strWkCmt) Then
                                        objRecMCmt = CType(objCmtMHash.Item(strWkCmt), Rec_MCmt)

                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If
                                    End If

                                End If

                            Next

                        End If

                    End If

                    ''禁止コメント
                    If Not objRecOrderSTD.CmtKinCdL Is Nothing AndAlso objRecOrderSTD.CmtKinCdL.Count > 0 Then

                        For Each strWkCmt As String In objRecOrderSTD.CmtKinCdL

                            strWkCmt = strWkCmt.Trim

                            If objCmtKinMHash.ContainsKey(strWkCmt) Then
                                objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If
                                End If

                            End If


                        Next

                    End If

                Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                    ''夕
                    If blnDelimiterKin = False Then

                        If Not objRecOrderSTD.CmtCdE Is Nothing AndAlso objRecOrderSTD.CmtCdE.Count > 0 Then

                            For Each strWkCmt As String In objRecOrderSTD.CmtCdE

                                strWkCmt = strWkCmt.Trim

                                If objCmtKinMHash.ContainsKey(strWkCmt) Then

                                    objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                        If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                            objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                        End If
                                    End If

                                Else

                                    If objCmtMHash.ContainsKey(strWkCmt) Then

                                        objRecMCmt = CType(objCmtMHash.Item(strWkCmt), Rec_MCmt)

                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                            If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                                objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                            End If
                                        End If

                                    End If

                                End If

                            Next

                        End If

                    End If

                    ''禁止コメント
                    If Not objRecOrderSTD.CmtKinCdE Is Nothing AndAlso objRecOrderSTD.CmtKinCdE.Count > 0 Then

                        For Each strWkCmt As String In objRecOrderSTD.CmtKinCdE

                            strWkCmt = strWkCmt.Trim

                            If objCmtKinMHash.ContainsKey(strWkCmt) Then

                                objRecMCmt = CType(objCmtKinMHash.Item(strWkCmt), Rec_MCmt)

                                If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                    If objSortOrderCmt.ContainsKey(strWkCmt.PadRight(8)) = False Then
                                        objSortOrderCmt.Add(strWkCmt, strWkCmt)
                                    End If
                                End If

                            End If

                        Next

                    End If

            End Select

            ''コメントをコード順にソートする
            Dim strWk As String

            For Each dicCmt As DictionaryEntry In objSortOrderCmt

                strWk = CType(dicCmt.Value, String)
                objOrderCmt.Add(strWk.PadRight(8))

            Next

            Return objOrderCmt

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    '**********************************************************************
    ' Name       :オーダ連携食種・主食変換処理
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTX)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :2016/05/06
    '        By  :FTC)T.IKEDA
    ' Comment    :食種によるコメント付加処理で、コメント２つ目と３つめの朝昼夕の区分チェック
    '             漏れから、不要なコメントが付加される障害対応
    '             朝昼夕うでのIF文チェックを追加した
    '**********************************************************************
    Private Function pfChangOChgMainDetail(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                           ByRef objRecOMain As Rec_OMain, _
                                           ByVal objOChgMain As SortedList, _
                                           ByVal strAgeCode As String, _
                                           ByVal strEatTmDiv As String, _
                                           ByVal blnDelimiterKin As Boolean, _
                                           ByVal objCmtKinMHash As Hashtable, _
                                           ByVal objCmtMHash As Hashtable, _
                                           ByVal objAllCmtAjiHash As Hashtable, _
                                           ByVal blnFindMEAL As Boolean, _
                                           ByVal blnFindMACK As Boolean, _
                                           ByVal blnFindAGE As Boolean, _
                                           ByVal blnFindEATTM As Boolean, _
                                           ByVal intFindCMT As Integer, _
                                           ByVal blnGetMeal As Boolean, _
                                           ByVal blnGetMainCk As Boolean, _
                                           ByVal blnGetAddCmt As Boolean _
                                           ) As Integer
        Try

            Dim intRtn As Integer = 0

            Dim strFindKey As String = ""       ''検索キー
            Dim blnFinded As Boolean = False    ''検索結果

            Dim objCmtOrdered As New StringCollection       ''コメントマスタ
            Dim objCmtKey As New StringCollection           ''コメント検索キー


            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
            strFindKey = ""

            ''オーダされたコメントを並び変える
            objCmtOrdered = pfSortOrderCmt(objRecOrderSTD, _
                                           blnDelimiterKin, _
                                           objCmtKinMHash, _
                                           objCmtMHash, _
                                           strEatTmDiv)

            ''時間ごとに処理する
            Select Case strEatTmDiv

                Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                    If blnFindMEAL = True Then
                        ''食種を探す
                        If objRecOrderSTD.MealCdM <> "" Then
                            strFindKey &= objRecOrderSTD.MealCdM.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MealCd <> "" Then
                            strFindKey &= objRecOrderSTD.MealCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                    If blnFindMACK = True Then
                        ''主食を探す
                        If objRecOrderSTD.MaCkCdM <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCdM.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MaCkCd <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If


                Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                    If blnFindMEAL = True Then
                        ''食種を探す
                        If objRecOrderSTD.MealCdL <> "" Then
                            strFindKey &= objRecOrderSTD.MealCdL.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MealCd <> "" Then
                            strFindKey &= objRecOrderSTD.MealCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                    If blnFindMACK = True Then
                        ''主食を探す
                        If objRecOrderSTD.MaCkCdL <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCdL.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MaCkCd <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                    If blnFindMEAL = True Then
                        ''食種を探す
                        If objRecOrderSTD.MealCdE <> "" Then
                            strFindKey &= objRecOrderSTD.MealCdE.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MealCd <> "" Then
                            strFindKey &= objRecOrderSTD.MealCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                    ''主食を探す
                    If blnFindMACK = True Then
                        If objRecOrderSTD.MaCkCdE <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCdE.Trim.PadRight(5)
                        ElseIf objRecOrderSTD.MaCkCd <> "" Then
                            strFindKey &= objRecOrderSTD.MaCkCd.Trim.PadRight(5)
                        Else
                            strFindKey &= "".PadRight(5)
                        End If
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                Case Else

                    ''食種
                    If blnFindMEAL = True Then
                        strFindKey &= objRecOrderSTD.MealCd.Trim.PadRight(5)
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

                    If blnFindMACK = True Then
                        ''主食
                        strFindKey &= objRecOrderSTD.MaCkCd.Trim.PadRight(5)
                    Else
                        strFindKey &= "".PadRight(5)
                    End If

            End Select

            ''年齢コード
            If blnFindAGE = True Then
                strFindKey &= strAgeCode.Trim.PadRight(3)
            Else
                strFindKey &= "".PadRight(3)
            End If

            ''時間区分
            If blnFindEATTM = True Then
                strFindKey &= strEatTmDiv.PadRight(2)
            Else
                strFindKey &= CType(ClsCst.ORecv.OCHG_OrderEatTmCnvInt, String).PadRight(2)
            End If

            ''コメントがある場合
            If Not objCmtOrdered Is Nothing AndAlso objCmtOrdered.Count > 0 _
               AndAlso intFindCMT > 0 Then

                ''コメントをループ
                For intLoopCnt As Integer = intFindCMT To 1 Step -1

                    Dim strFindKey2 As String = ""
                    strFindKey2 = strFindKey

                    If objCmtOrdered.Count >= intLoopCnt Then

                        ''コメントの全パターンを取得
                        objCmtKey = CommentCombinationCollection(objCmtOrdered.Count, intLoopCnt, objCmtOrdered)

                        For Each strCmtKey As String In objCmtKey

                            ''コメントキー
                            strFindKey2 = strFindKey

                            strFindKey2 &= strCmtKey

                            If intLoopCnt = 1 Then
                                strFindKey2 &= "".PadRight(8) & "".PadRight(8)
                            ElseIf intLoopCnt = 2 Then
                                strFindKey2 &= "".PadRight(8)
                            End If

                            If objOChgMain.ContainsKey(strFindKey2) Then

                                ''見つけた場合
                                Dim objRecOChg As New Rec_OChg
                                objRecOChg = CType(objOChgMain.Item(strFindKey2), Rec_OChg)

                                Dim objRecMCmt As New Rec_MCmt

                                Select Case strEatTmDiv

                                    Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                        ''食種を取得するなら
                                        If blnGetMeal = True Then
                                            objRecOMain.MealCdMNew = objRecOChg.MealCd      ''食種
                                        End If

                                        ''主食を取得するなら
                                        If blnGetMainCk = True Then
                                            objRecOMain.MaCkCdMNew = objRecOChg.MaCkCd      ''主食
                                        End If

                                        ''コメントを追加するなら
                                        If blnGetAddCmt = True Then

                                            ''コメント
                                            objRecOMain.CmtAddedMNew = True

                                            If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                        objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd1.Trim
                                                    End If
                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddMNew = "" Then
                                                            objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd2.Trim
                                                        Else
                                                            objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                        End If
                                                    End If
                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddMNew = "" Then
                                                            objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd3.Trim
                                                        Else
                                                            objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                        End If
                                                    End If
                                                End If

                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                        ''昼

                                        ''食種を取得するなら
                                        If blnGetMeal = True Then
                                            objRecOMain.MealCdLNew = objRecOChg.MealCd      ''食種
                                        End If

                                        ''主食を取得するなら
                                        If blnGetMainCk = True Then
                                            objRecOMain.MaCkCdLNew = objRecOChg.MaCkCd      ''主食
                                        End If

                                        ''コメントを追加するなら
                                        If blnGetAddCmt = True Then
                                            objRecOMain.CmtAddedLNew = True

                                            If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                        objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd1.Trim
                                                    End If

                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddLNew = "" Then
                                                            objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd2.Trim
                                                        Else
                                                            objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                        End If
                                                    End If

                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                                    '==========================================================
                                                    '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddLNew = "" Then
                                                            objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd3.Trim
                                                        Else
                                                            objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                        End If
                                                    End If
                                                    '==========================================================
                                                    '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                End If

                                            End If
                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                        ''夕
                                        ''食種を取得するなら
                                        If blnGetMeal = True Then
                                            objRecOMain.MealCdENew = objRecOChg.MealCd      ''食種
                                        End If

                                        ''主食を取得するなら
                                        If blnGetMainCk = True Then
                                            objRecOMain.MaCkCdENew = objRecOChg.MaCkCd      ''主食
                                        End If

                                        ''コメントを追加するなら
                                        If blnGetAddCmt = True Then
                                            objRecOMain.CmtAddedENew = True

                                            If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                        objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd1.Trim
                                                    End If

                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                                    '==========================================================
                                                    '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddENew = "" Then
                                                            objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd2.Trim
                                                        Else
                                                            objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                        End If
                                                    End If
                                                    '==========================================================
                                                    '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                End If

                                            End If

                                            If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                    objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                                    '==========================================================
                                                    '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                        If objRecOMain.CmtAddENew = "" Then
                                                            objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd3.Trim
                                                        Else
                                                            objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                        End If
                                                    End If
                                                    '==========================================================
                                                    '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                    '==========================================================
                                                End If
                                            End If
                                        End If

                                    Case Else

                                            ''朝・昼・夕以外
                                            ''食種を取得するなら
                                            If blnGetMeal = True Then
                                                If objRecOMain.MealCdMNew.Trim = "" Then
                                                    objRecOMain.MealCdMNew = objRecOChg.MealCd      ''食種
                                                End If

                                                If objRecOMain.MealCdLNew.Trim = "" Then
                                                    objRecOMain.MealCdLNew = objRecOChg.MealCd      ''食種
                                                End If

                                                If objRecOMain.MealCdENew.Trim = "" Then
                                                    objRecOMain.MealCdENew = objRecOChg.MealCd      ''食種
                                                End If
                                            End If

                                            ''主食を取得するなら
                                            If blnGetMainCk = True Then
                                                If objRecOMain.MaCkCdMNew.Trim = "" Then
                                                    objRecOMain.MaCkCdMNew = objRecOChg.MaCkCd      ''主食
                                                End If

                                                If objRecOMain.MaCkCdLNew.Trim = "" Then
                                                    objRecOMain.MaCkCdLNew = objRecOChg.MaCkCd      ''主食
                                                End If

                                                If objRecOMain.MaCkCdENew.Trim = "" Then
                                                    objRecOMain.MaCkCdENew = objRecOChg.MaCkCd      ''主食
                                                End If
                                            End If

                                            ''コメントを追加するなら
                                            If blnGetAddCmt = True Then
                                                ''コメント（朝）
                                                If objRecOMain.CmtAddedMNew = False Then

                                                    If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                                objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd1.Trim
                                                            End If
                                                        End If

                                                    End If

                                                    If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                                If objRecOMain.CmtAddMNew = "" Then
                                                                    objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd2.Trim
                                                                Else
                                                                    objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                                End If
                                                            End If
                                                        End If

                                                    End If

                                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)

                                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                                If objRecOMain.CmtAddMNew = "" Then
                                                                    objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd3.Trim
                                                                Else
                                                                    objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                                End If
                                                            End If
                                                        End If

                                                    End If

                                                    objRecOMain.CmtAddedMNew = True

                                                End If

                                                ''昼
                                                If objRecOMain.CmtAddedLNew = False Then

                                                    If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                                objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd1.Trim
                                                            End If

                                                        End If

                                                    End If

                                                    If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                                If objRecOMain.CmtAddLNew = "" Then
                                                                    objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd2.Trim
                                                                Else
                                                                    objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                                End If
                                                            End If

                                                        End If

                                                    End If

                                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                                        '==========================================================
                                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                            If objRecOMain.CmtAddLNew = "" Then
                                                                objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd3.Trim
                                                            Else
                                                                objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                            End If
                                                        End If
                                                        '==========================================================
                                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                    End If

                                                End If

                                                objRecOMain.CmtAddedLNew = True

                                            End If

                                            ''夕
                                            If objRecOMain.CmtAddedENew = False Then

                                                If objRecOChg.AddCmtCd1.Trim <> "" Then

                                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                            objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd1.Trim
                                                        End If

                                                    End If

                                                End If

                                                If objRecOChg.AddCmtCd2.Trim <> "" Then

                                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)
                                                        '==========================================================
                                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                            If objRecOMain.CmtAddENew = "" Then
                                                                objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd2.Trim
                                                            Else
                                                                objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                            End If
                                                        End If
                                                        '==========================================================
                                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                    End If

                                                    End If

                                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                                        '==========================================================
                                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                            If objRecOMain.CmtAddENew = "" Then
                                                                objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd3.Trim
                                                            Else
                                                                objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                            End If
                                                        End If
                                                        '==========================================================
                                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                                        '==========================================================
                                                    End If

                                                End If

                                                objRecOMain.CmtAddedENew = True

                                            End If
                                        End If
                                End Select

                                blnFinded = True
                                Exit For

                            End If

                        Next

                    End If

                    ''見つけた瞬間ループ終了
                    If blnFinded = True Then
                        Exit For
                    End If

                Next

            End If

            ''見つからない場合
            If blnFinded = False Then

                ''コメントなしとみなして
                strFindKey &= "".PadRight(8)
                strFindKey &= "".PadRight(8)
                strFindKey &= "".PadRight(8)

                If objOChgMain.ContainsKey(strFindKey) Then

                    ''見つけた場合
                    Dim objRecOChg As New Rec_OChg
                    objRecOChg = CType(objOChgMain.Item(strFindKey), Rec_OChg)

                    Dim objRecMCmt As New Rec_MCmt

                    Select Case strEatTmDiv

                        Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                            ''朝
                            ''食種を取得するなら
                            If blnGetMeal = True Then
                                objRecOMain.MealCdMNew = objRecOChg.MealCd      ''食種
                            End If

                            ''主食を取得するなら
                            If blnGetMainCk = True Then
                                objRecOMain.MaCkCdMNew = objRecOChg.MaCkCd      ''主食
                            End If

                            ''コメントを追加するなら
                            If blnGetAddCmt = True Then
                                ''コメント
                                objRecOMain.CmtAddedMNew = True

                                If objRecOChg.AddCmtCd1.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                            objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd1.Trim
                                        End If
                                    End If

                                End If

                                If objRecOChg.AddCmtCd2.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddMNew = "" Then
                                                objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd2.Trim
                                            Else
                                                objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                            End If
                                        End If
                                    End If

                                End If

                                If objRecOChg.AddCmtCd3.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddMNew = "" Then
                                                objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd3.Trim
                                            Else
                                                objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                            End If
                                        End If
                                    End If

                                End If
                            End If

                        Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                            ''昼
                            ''食種を取得するなら
                            If blnGetMeal = True Then
                                objRecOMain.MealCdLNew = objRecOChg.MealCd      ''食種
                            End If

                            ''主食を取得するなら
                            If blnGetMainCk = True Then
                                objRecOMain.MaCkCdLNew = objRecOChg.MaCkCd      ''主食
                            End If

                            ''コメントを追加するなら
                            If blnGetAddCmt = True Then

                                objRecOMain.CmtAddedLNew = True

                                If objRecOChg.AddCmtCd1.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                            objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd1.Trim
                                        End If

                                    End If

                                End If

                                If objRecOChg.AddCmtCd2.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddLNew = "" Then
                                                objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd2.Trim
                                            Else
                                                objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                            End If
                                        End If

                                    End If

                                End If

                                If objRecOChg.AddCmtCd3.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                        '==========================================================
                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                        If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddLNew = "" Then
                                                objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd3.Trim
                                            Else
                                                objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                            End If
                                        End If
                                        '==========================================================
                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                    End If

                                End If
                            End If

                        Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                            ''夕
                            ''食種を取得するなら
                            If blnGetMeal = True Then
                                objRecOMain.MealCdENew = objRecOChg.MealCd      ''食種
                            End If

                            ''主食を取得するなら
                            If blnGetMainCk = True Then
                                objRecOMain.MaCkCdENew = objRecOChg.MaCkCd      ''主食
                            End If

                            ''コメントを追加するなら
                            If blnGetAddCmt = True Then

                                objRecOMain.CmtAddedENew = True

                                If objRecOChg.AddCmtCd1.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                            objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd1.Trim
                                        End If

                                    End If

                                End If

                                If objRecOChg.AddCmtCd2.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)
                                        '==========================================================
                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddENew = "" Then
                                                objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd2.Trim
                                            Else
                                                objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                            End If
                                        End If
                                        '==========================================================
                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                    End If

                                End If

                                If objRecOChg.AddCmtCd3.Trim <> "" Then

                                    If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                        objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                        '==========================================================
                                        '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                        If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                            If objRecOMain.CmtAddENew = "" Then
                                                objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd3.Trim
                                            Else
                                                objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                            End If
                                        End If
                                        '==========================================================
                                        '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                        '==========================================================
                                    End If

                                End If
                            End If

                        Case Else

                            ''食種を取得するなら
                            If blnGetMeal = True Then
                                ''朝・昼・夕以外
                                If objRecOMain.MealCdMNew.Trim = "" Then
                                    objRecOMain.MealCdMNew = objRecOChg.MealCd      ''食種
                                End If

                                If objRecOMain.MealCdLNew.Trim = "" Then
                                    objRecOMain.MealCdLNew = objRecOChg.MealCd      ''食種
                                End If

                                If objRecOMain.MealCdENew.Trim = "" Then
                                    objRecOMain.MealCdENew = objRecOChg.MealCd      ''食種
                                End If
                            End If

                            ''主食を取得するなら
                            If blnGetMainCk = True Then
                                If objRecOMain.MaCkCdMNew.Trim = "" Then
                                    objRecOMain.MaCkCdMNew = objRecOChg.MaCkCd      ''主食
                                End If

                                If objRecOMain.MaCkCdLNew.Trim = "" Then
                                    objRecOMain.MaCkCdLNew = objRecOChg.MaCkCd      ''主食
                                End If

                                If objRecOMain.MaCkCdENew.Trim = "" Then
                                    objRecOMain.MaCkCdENew = objRecOChg.MaCkCd      ''主食
                                End If
                            End If

                            ''コメントを追加するなら
                            If blnGetAddCmt = True Then
                                ''コメント（朝）
                                If objRecOMain.CmtAddedMNew = False Then

                                    If objRecOChg.AddCmtCd1.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd1.Trim
                                            End If
                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd2.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddMNew = "" Then
                                                    objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd2.Trim
                                                Else
                                                    objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                End If
                                            End If
                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddMNew = "" Then
                                                    objRecOMain.CmtAddMNew &= objRecOChg.AddCmtCd3.Trim
                                                Else
                                                    objRecOMain.CmtAddMNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                End If
                                            End If
                                        End If

                                    End If

                                    objRecOMain.CmtAddedMNew = True

                                End If

                                ''昼
                                If objRecOMain.CmtAddedLNew = False Then

                                    If objRecOChg.AddCmtCd1.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd1.Trim
                                            End If

                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd2.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddLNew = "" Then
                                                    objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd2.Trim
                                                Else
                                                    objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                End If
                                            End If

                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                            '==========================================================
                                            '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddLNew = "" Then
                                                    objRecOMain.CmtAddLNew &= objRecOChg.AddCmtCd3.Trim
                                                Else
                                                    objRecOMain.CmtAddLNew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                End If
                                            End If
                                            '==========================================================
                                            '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                        End If

                                    End If

                                    objRecOMain.CmtAddedLNew = True

                                End If

                                ''夕
                                If objRecOMain.CmtAddedENew = False Then

                                    If objRecOChg.AddCmtCd1.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd1.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd1.Trim), Rec_MCmt)

                                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd1.Trim
                                            End If

                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd2.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd2.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd2.Trim), Rec_MCmt)
                                            '==========================================================
                                            '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddENew = "" Then
                                                    objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd2.Trim
                                                Else
                                                    objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd2.Trim
                                                End If
                                            End If
                                            '==========================================================
                                            '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                        End If

                                    End If

                                    If objRecOChg.AddCmtCd3.Trim <> "" Then

                                        If objAllCmtAjiHash.ContainsKey(objRecOChg.AddCmtCd3.Trim) Then

                                            objRecMCmt = CType(objAllCmtAjiHash.Item(objRecOChg.AddCmtCd3.Trim), Rec_MCmt)
                                            '==========================================================
                                            '↓　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                                If objRecOMain.CmtAddENew = "" Then
                                                    objRecOMain.CmtAddENew &= objRecOChg.AddCmtCd3.Trim
                                                Else
                                                    objRecOMain.CmtAddENew &= ORDER_DELIMITER_CHAR_PIPE & objRecOChg.AddCmtCd3.Trim
                                                End If
                                            End If
                                            '==========================================================
                                            '↑　2016/05/06　T.IKEDA)区分チェックを追加
                                            '==========================================================
                                        End If

                                        End If

                                        objRecOMain.CmtAddedENew = True

                                    End If
                                End If

                    End Select

                    blnFinded = True

                End If

            End If

            ''データを見つけた場合、1を返す
            If blnFinded = True Then
                intRtn = 1
            End If

            Return intRtn

        Catch ex As Exception

            Return -1

        End Try

    End Function

    '**********************************************************************
    ' Name       :オーダ連携食種・主食変換処理
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
    Private Function pfChangOChgMain(ByVal strOrderChangeCd As String, _
                                     ByVal objRecOrderSTD As Rec_OrderSTD, _
                                     ByRef objRecOMain As Rec_OMain, _
                                     ByVal blnDelimiterKin As Boolean, _
                                     ByVal objCmtKinMHash As Hashtable, _
                                     ByVal objCmtMHash As Hashtable, _
                                     ByVal objAllCmtAjiHash As Hashtable, _
                                     ByVal objOChgAge As SortedList, _
                                     ByVal objOChgMain As SortedList, _
                                     ByVal objRecORecv As Rec_ORecv, _
                            Optional ByVal blnGetMeal As Boolean = True, _
                            Optional ByVal blnGetMainCk As Boolean = True, _
                            Optional ByVal blnGetAddCmt As Boolean = True _
                                     ) As Boolean

        Try

            Dim strAgeCode As String = ""                   ''年齢コード

            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
            '変換処理
            Dim blnFindMEALMain As Boolean
            Dim blnFindMACKMain As Boolean
            Dim blnFindAGEMain As Boolean
            Dim blnFindEATTMMain As Boolean

            Dim intFindMEAL As Integer = 0
            Dim intFindMACK As Integer = 0
            Dim intFindAGE As Integer = 0
            Dim intFindEATTM As Integer = 0

            Dim intFindCMT As Integer

            ''食種変換するかどうか
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderMealCnvSet) = -1 Then
                blnFindMEALMain = False
                intFindMEAL = 1
            Else
                blnFindMEALMain = True
                intFindMEAL = 2
            End If

            ''主食変換するかどうか
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderMaCkCnvSet) = -1 Then
                blnFindMACKMain = False
                intFindMACK = 1
            Else
                blnFindMACKMain = True
                intFindMACK = 2
            End If

            ''年齢変換するかどうか
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderAgeCnvSet) = -1 Then
                blnFindAGEMain = False
                intFindAGE = 1
            Else
                blnFindAGEMain = True
                intFindAGE = 2

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                '年齢情報を取得
                Dim intAge As Integer = 0                       ''年齢
                Dim intSex As Integer = 0                       ''性別

                ''年齢を取得
                If objRecOrderSTD.Birth <> "" Then
                    intAge = GetPerAgeInt(objRecOrderSTD.Birth)
                    If intAge = -1 Then
                        intAge = GetPerAgeInt(objRecORecv.Birth)
                    End If
                Else
                    intAge = GetPerAgeInt(objRecORecv.Birth)
                End If

                ''性別を取得
                intSex = objRecOrderSTD.Sex

                If intSex = 0 Then
                    intSex = objRecORecv.Sex
                End If

                ''年齢コードを取得
                If intAge <> -1 AndAlso intSex <> 0 Then

                    If Not objOChgAge Is Nothing Then

                        If objOChgAge.ContainsKey(intSex.ToString & intAge.ToString("000")) = True Then
                            strAgeCode = CType(objOChgAge(intSex.ToString & intAge.ToString("000")), String)
                        Else
                            strAgeCode = ""
                        End If
                    Else
                        strAgeCode = ""
                    End If

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''コメント変換
            intFindCMT = 0
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderCmt1CnvSet) <> -1 Then
                intFindCMT = 1
            End If
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderCmt2CnvSet) <> -1 Then
                intFindCMT = 2
            End If
            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderCmt3CnvSet) <> -1 Then
                intFindCMT = 3
            End If

            ''時間帯ごとにループする
            Dim objEatTmDiv As New StringCollection

            If strOrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderEatTmCnvSet) = -1 Then

                ''時間変換なし
                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.MorningDiv, String))
                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.LunchDiv, String))
                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.EveningDiv, String))
                blnFindEATTMMain = False
                intFindEATTM = 1

            Else

                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.MorningDiv, String))
                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.LunchDiv, String))
                objEatTmDiv.Add(CType(ClsCst.enmEatTimeDiv.EveningDiv, String))
                blnFindEATTMMain = True
                intFindEATTM = 2

            End If

            ''見つけたかどうか
            Dim intFinded As Integer
            Dim blnFindedM As Boolean = False
            Dim blnFindedL As Boolean = False
            Dim blnFindedE As Boolean = False

            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
            ''食種キー構成
            For intLoopMEAL As Integer = 1 To intFindMEAL

                Dim blnFindMEAL As Boolean = True

                If blnFindMEALMain = False Then
                    blnFindMEAL = False
                Else
                    If intLoopMEAL = 2 Then
                        blnFindMEAL = False
                    End If
                End If

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                ''主食キー構成
                For intLoopMACK As Integer = 1 To intFindMACK

                    Dim blnFindMACK As Boolean = True

                    If blnFindMACKMain = False Then
                        blnFindMACK = False
                    Else
                        If intLoopMACK = 2 Then
                            blnFindMACK = False
                        End If
                    End If


                    '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                    ''年齢キー構成
                    For intLoopAGE As Integer = 1 To intFindAGE

                        Dim blnFindAGE As Boolean = True

                        If blnFindAGEMain = False Then
                            blnFindAGE = False
                        Else
                            If intLoopAGE = 2 Then
                                blnFindAGE = False
                            End If
                        End If

                        '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                        For intLoopEatTm As Integer = 1 To intFindEATTM

                            Dim blnFindEATTM As Boolean = True

                            If blnFindEATTMMain = False Then
                                blnFindEATTM = False
                            Else
                                If intLoopEatTm = 2 Then
                                    blnFindEATTM = False
                                End If
                            End If


                            '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                            ''時間ごとループ
                            For Each strEatTmDiv As String In objEatTmDiv

                                intFinded = 0

                                Select Case strEatTmDiv

                                    Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)

                                        If blnFindedM = False Then

                                            intFinded = pfChangOChgMainDetail(objRecOrderSTD, _
                                                                              objRecOMain, _
                                                                              objOChgMain, _
                                                                              strAgeCode, _
                                                                              strEatTmDiv, _
                                                                              blnDelimiterKin, _
                                                                              objCmtKinMHash, _
                                                                              objCmtMHash, _
                                                                              objAllCmtAjiHash, _
                                                                              blnFindMEAL, _
                                                                              blnFindMACK, _
                                                                              blnFindAGE, _
                                                                              blnFindEATTM, _
                                                                              intFindCMT, _
                                                                              blnGetMeal, _
                                                                              blnGetMainCk, _
                                                                              blnGetAddCmt _
                                                                              )

                                        End If

                                    Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)

                                        If blnFindedL = False Then

                                            intFinded = pfChangOChgMainDetail(objRecOrderSTD, _
                                                                              objRecOMain, _
                                                                              objOChgMain, _
                                                                              strAgeCode, _
                                                                              strEatTmDiv, _
                                                                              blnDelimiterKin, _
                                                                              objCmtKinMHash, _
                                                                              objCmtMHash, _
                                                                              objAllCmtAjiHash, _
                                                                              blnFindMEAL, _
                                                                              blnFindMACK, _
                                                                              blnFindAGE, _
                                                                              blnFindEATTM, _
                                                                              intFindCMT, _
                                                                              blnGetMeal, _
                                                                              blnGetMainCk, _
                                                                              blnGetAddCmt _
                                                                              )
                                        End If


                                    Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)

                                        If blnFindedE = False Then

                                            intFinded = pfChangOChgMainDetail(objRecOrderSTD, _
                                                                              objRecOMain, _
                                                                              objOChgMain, _
                                                                              strAgeCode, _
                                                                              strEatTmDiv, _
                                                                              blnDelimiterKin, _
                                                                              objCmtKinMHash, _
                                                                              objCmtMHash, _
                                                                              objAllCmtAjiHash, _
                                                                              blnFindMEAL, _
                                                                              blnFindMACK, _
                                                                              blnFindAGE, _
                                                                              blnFindEATTM, _
                                                                              intFindCMT, _
                                                                              blnGetMeal, _
                                                                              blnGetMainCk, _
                                                                              blnGetAddCmt _
                                                                              )
                                        End If


                                End Select

                                If intFinded = -1 Then

                                    Throw New Exception("変換処理でエラーが発生しました。")

                                Else

                                    If intFinded = 1 Then

                                        Select Case strEatTmDiv

                                            Case CType(ClsCst.enmEatTimeDiv.MorningDiv, String)
                                                blnFindedM = True
                                            Case CType(ClsCst.enmEatTimeDiv.LunchDiv, String)
                                                blnFindedL = True
                                            Case CType(ClsCst.enmEatTimeDiv.EveningDiv, String)
                                                blnFindedE = True
                                        End Select

                                    End If

                                End If

                                ''データを全部見つけた場合
                                If blnFindedM = True AndAlso blnFindedL = True AndAlso blnFindedE = True Then
                                    Exit For
                                End If

                            Next
                            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                            ''データを全部見つけた場合
                            If blnFindedM = True AndAlso blnFindedL = True AndAlso blnFindedE = True Then
                                Exit For
                            End If

                        Next
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                        ''データを全部見つけた場合
                        If blnFindedM = True AndAlso blnFindedL = True AndAlso blnFindedE = True Then
                            Exit For
                        End If

                    Next
                    '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    ''データを全部見つけた場合
                    If blnFindedM = True AndAlso blnFindedL = True AndAlso blnFindedE = True Then
                        Exit For
                    End If

                Next
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                ''データを全部見つけた場合
                If blnFindedM = True AndAlso blnFindedL = True AndAlso blnFindedE = True Then
                    Exit For
                End If

            Next
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／


            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    '**********************************************************************
    ' Name       :オーダ連携共通Omain変換処理
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
    Public Function OrderDataChangeToOmain(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                           ByVal objExecDivHash As SortedList, _
                                           ByVal objOdDivHash As SortedList, _
                                           ByVal objMealMHash As Hashtable, _
                                           ByVal objCmtNorMHash As Hashtable, _
                                           ByVal objCmtKinMHash As Hashtable, _
                                           ByVal SysInf As Rec_XSys, _
                                           ByVal objBumonCmtNorMHash As Hashtable, _
                                           ByVal objBumonCmtKinMHash As Hashtable, _
                                           ByVal objAllCmtAjiHash As Hashtable, _
                                           ByVal objOChgAge As SortedList, _
                                           ByVal objOChgMain As SortedList, _
                                           ByVal objRecORecv As Rec_ORecv) As Rec_OMain

        Try

            Dim objRecOMain As New Rec_OMain

            Dim blnSnack10 As Boolean
            Dim blnSnack15 As Boolean
            Dim blnSnack20 As Boolean

            Dim blnDelimiterCmtKin As Boolean

            ''コメントマスタ別々なら、TRUE
            If SysInf.OrderDelimiterCmtKin = 0 Then
                blnDelimiterCmtKin = False
            Else
                blnDelimiterCmtKin = True
            End If

            ''食種・主食変換テーブルを使用する場合            
            If SysInf.OrderChangeCd.Trim <> "" Then

                ''食種・主食変換テーブルを使用する場合
                If pfChangOChgMain(SysInf.OrderChangeCd.Trim, objRecOrderSTD, objRecOMain, blnDelimiterCmtKin, _
                                objCmtKinMHash, objCmtNorMHash, objAllCmtAjiHash, objOChgAge, objOChgMain, objRecORecv) = False Then

                    ''エラー時

                End If

            Else

                ''食種のみ変換
                If SysInf.OrderChangeCdMealCmt.Trim <> "" Then

                    ''変換テーブルを使用し、食種・コメントを取得する場合()
                    If pfChangOChgMain(SysInf.OrderChangeCdMealCmt.Trim, objRecOrderSTD, objRecOMain, blnDelimiterCmtKin, _
                                       objCmtKinMHash, objCmtNorMHash, objAllCmtAjiHash, objOChgAge, objOChgMain, objRecORecv, True, False, True) = False Then
                        ''エラー時
                    End If

                End If

                ''主食のみ変換
                If SysInf.OrderChangeCdMainCok.Trim <> "" Then
                    ''変換テーブルを使用し、主食を取得する場合()
                    If pfChangOChgMain(SysInf.OrderChangeCdMainCok.Trim, objRecOrderSTD, objRecOMain, blnDelimiterCmtKin, _
                                       objCmtKinMHash, objCmtNorMHash, objAllCmtAjiHash, objOChgAge, objOChgMain, objRecORecv, False, True, False) = False Then
                        ''エラー時
                    End If

                End If

            End If

            With objRecOMain

                .OpYMD = objRecOrderSTD.OpYMD                     '登録年月日()
                .OpNo = objRecOrderSTD.OpNo                       '登録番号()
                .OdTp = objRecOrderSTD.OdTp                       'オーダ種別()
                .RfctDiv = objRecOrderSTD.RfctDiv                 '反映区分()
                .PpIsDiv = objRecOrderSTD.PpIsDiv                 '食事箋発行区分()
                .OdNo = objRecOrderSTD.OdNo                       'オーダ番号()
                .Rsheet = objRecOrderSTD.Rsheet                   '食事箋/伝票コード
                .AtvOd = objRecOrderSTD.AtvOd                     '実施オーダ区分()
                .Deadline = objRecOrderSTD.Deadline               '締切り区分()
                .ExecDiv = objRecOrderSTD.ExecDiv                 '処理区分()
                .TelMDiv = objRecOrderSTD.TelMDiv                 '電文種別()
                .OdDiv = objRecOrderSTD.OdDiv                     'オーダ指示区分()

                '共通処理区分
                .ComExecDiv = pfChangComExecDiv(objRecOrderSTD, objExecDivHash)

                .ComOdDiv = pfChangComOdDiv(objRecOrderSTD, objOdDivHash)   '共通指示区分 

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                If objRecOrderSTD.Snack = 1 Then

                    blnSnack10 = True
                    blnSnack15 = True
                    blnSnack20 = True

                ElseIf objRecOrderSTD.Snack = 0 Then

                    blnSnack10 = False
                    blnSnack15 = False
                    blnSnack20 = False

                Else
                    ''間食があるかどうか決定する
                    If objRecOrderSTD.Snack10 = 1 Then
                        blnSnack10 = True
                    Else
                        blnSnack10 = False
                    End If

                    If objRecOrderSTD.Snack15 = 1 Then
                        blnSnack15 = True
                    Else
                        blnSnack15 = False
                    End If

                    If objRecOrderSTD.Snack20 = 1 Then
                        blnSnack20 = True
                    Else
                        blnSnack20 = False
                    End If

                End If

                .PerCd = objRecOrderSTD.PerCd       ''個人コード

                .StYMD = objRecOrderSTD.StYMD       ''開始年月日
                .StTmDiv = objRecOrderSTD.StTmDiv   ''開始時間区分

                If objRecOrderSTD.EdYMD.Trim = "" Then
                    .EdYMD = ClsCst.PerCST.LIMITVALUE_T_PerYMD
                    .EdTmDiv = ClsCst.PerCST.LIMITVALUE_T_PerTmDiv
                Else

                    .EdYMD = objRecOrderSTD.EdYMD.Trim
                    If objRecOrderSTD.EdTmDiv = 0 Then
                        .EdTmDiv = ClsCst.enmEatTimeDiv.NightDiv
                    Else
                        .EdTmDiv = objRecOrderSTD.EdTmDiv
                    End If
                End If

                .MovYMD = objRecOrderSTD.MovYMD         ''移動年月日
                .MovTime = objRecOrderSTD.MovTime       ''移動時間

                ''開始日がない場合は移動年月日を見る
                If .StYMD = "" OrElse .StYMD = ClsCst.PerCST.MINVALUE_T_PerYMD Then

                    If .MovYMD = "" OrElse .MovYMD = ClsCst.PerCST.MINVALUE_T_PerYMD Then

                        .StYMD = .OpYMD
                        .StTmDiv = ChangeTimeToTmDiv(Now.ToString("HHmm"), SysInf)
                    Else

                        .StYMD = .MovYMD
                        .StTmDiv = ChangeTimeToTmDiv(.MovTime, SysInf)

                    End If

                End If

                .PerNm = objRecOrderSTD.PerNm.Trim           ''個人名
                .PerKana = objRecOrderSTD.PerKana.Trim       ''個人名カナ
                .Sex = objRecOrderSTD.Sex                    ''性別

                .Birth = objRecOrderSTD.Birth           ''生年月日

                ''個人備考
                Dim objRefColM As New StringCollection
                Dim strSplitVal As String = ""

                Select Case SysInf.OrderPerRefBreak

                    Case "vbTab"

                        strSplitVal = vbTab

                    Case "vbCrLf"

                        strSplitVal = vbCrLf

                    Case "vbCr"
                        strSplitVal = vbCr

                    Case "vbLf"
                        strSplitVal = vbLf

                    Case Else

                        strSplitVal = SysInf.OrderPerRefBreak

                End Select

                If strSplitVal <> "" Then

                    objRefColM = SplitOrderData(objRecOrderSTD.RefColM.Trim, strSplitVal)

                    ''データがあった場合
                    If Not objRefColM Is Nothing AndAlso objRefColM.Count > 0 Then

                        Dim i As Integer = 0

                        For Each strRefColM As String In objRefColM

                            If i = 0 Then
                                .RefColM = strRefColM
                            Else
                                .RefColM &= vbCrLf & strRefColM
                            End If

                            i = i + 1
                        Next

                    Else

                        .RefColM = objRecOrderSTD.RefColM.Trim

                    End If

                Else

                    .RefColM = objRecOrderSTD.RefColM.Trim  ''マスタ備考

                End If

                .PerHigh = objRecOrderSTD.PerHigh       ''身長
                .PerWeit = objRecOrderSTD.PerWeit       ''体重 

                ''身体活動レベル
                If objRecOrderSTD.LifeLv.Trim = "" Then
                    .LifeLv = SysInf.PerLifeLv
                Else
                    .LifeLv = objRecOrderSTD.LifeLv.Trim
                End If

                ''診療科
                .MedCare = pfChangMedCare(objRecOrderSTD.MedCare.Trim)

                .PDoctor = objRecOrderSTD.PDoctor.Trim       ''主治医
                .SDoctor = objRecOrderSTD.SDoctor.Trim       ''担当医
                .NstDiv = objRecOrderSTD.NstDiv             ''NstDiv

                ''個人フリーの備考
                If objRecOrderSTD.RefColF.Trim <> "" Then

                    .RefColFM = objRecOrderSTD.RefColF.Trim
                    .RefColF10 = objRecOrderSTD.RefColF.Trim
                    .RefColFL = objRecOrderSTD.RefColF.Trim
                    .RefColF15 = objRecOrderSTD.RefColF.Trim
                    .RefColFE = objRecOrderSTD.RefColF.Trim
                    .RefColF20 = objRecOrderSTD.RefColF.Trim

                Else

                    .RefColFM = objRecOrderSTD.RefColFM.Trim
                    .RefColF10 = objRecOrderSTD.RefColF10.Trim
                    .RefColFL = objRecOrderSTD.RefColFL.Trim
                    .RefColF15 = objRecOrderSTD.RefColF15.Trim
                    .RefColFE = objRecOrderSTD.RefColFE.Trim
                    .RefColF20 = objRecOrderSTD.RefColF20.Trim

                End If

                ''病棟
                .Ward = pfGetWardCd(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                ''病室
                .Room = pfGetRoodCd(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                ''ベッド
                .Bed = pfGetBed(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                ''有効フラグ
                .ValidF = pfGetValidF(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                ''食種コード
                Dim objHashMeal As SortedList
                objHashMeal = pfGetMealCd(objRecOrderSTD, objRecOMain, SysInf, blnSnack10, blnSnack15, blnSnack20)

                ''食種
                If Not objHashMeal Is Nothing AndAlso objHashMeal.Count > 0 Then

                    If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) = True Then

                        .MealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)

                    Else

                        For intCnt As Integer = 0 To objHashMeal.Count - 1

                            If intCnt = 0 Then
                                .MealCd = CType(objHashMeal.GetByIndex(intCnt), String)
                            Else
                                .MealCd &= ORDER_DELIMITER_CHAR_PIPE & CType(objHashMeal.GetByIndex(intCnt), String)
                            End If

                        Next

                    End If

                Else
                    .MealCd = ""
                End If

                ''欠食
                .NoMeal = pfGetNoMeal(objRecOrderSTD, objHashMeal, SysInf, .ComOdDiv, blnSnack10, blnSnack15, blnSnack20)

                ''欠食理由
                .NoMlRs = pfGetNoMlRs(objRecOrderSTD, objHashMeal, SysInf, .ComOdDiv, blnSnack10, blnSnack15, blnSnack20)

                ''主食コード
                .MaCkCd = pfGetMaCkCd(objRecOrderSTD, objRecOMain, objHashMeal, objMealMHash, SysInf.OrderCnvMainFdCd, blnSnack10, blnSnack15, blnSnack20)

                '形態
                .SdShap = pfGetSdShap(objRecOrderSTD, objHashMeal, objMealMHash, blnSnack10, blnSnack15, blnSnack20)

                '特食加算
                .EatAdDiv = pfGetEatAdDiv(objRecOrderSTD, objHashMeal, objMealMHash, blnSnack10, blnSnack15, blnSnack20)

                '食札
                .MCPrt = pfGetMCPrt(objRecOrderSTD, objHashMeal, objMealMHash, blnSnack10, blnSnack15, blnSnack20)

                '選択食区分（デフォルト）
                .SelEtDivDefM = pfGetSelEtDivDefM(objRecOrderSTD, SysInf, blnSnack10, blnSnack15, blnSnack20)

                '配膳先
                .DelivCd = pfGetDelivCd(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                '選択食
                .SelEtDiv = pfGetSelEtDiv(objRecOrderSTD)

                ''食数の変換
                .KCnt = pfGetKCnt(objRecOrderSTD, blnSnack10, blnSnack15, blnSnack20)

                '／￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣￣＼
                ''アレルギーコメントがあるかどうかチェックをする
                If SysInf.OrderAllergyCmt <> "" Then

                    Dim blnAllergyCmt As Boolean

                    blnAllergyCmt = pfCheckAllergyCmt(SysInf, objRecOrderSTD)

                    ''アレルギーがある場合
                    If blnAllergyCmt = True Then

                        .OpNoRef = ClsCst.ORecv.OrderRsheetNm_Allergy

                    End If

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                '共通コメントコード
                .CmtCd = pfGetCmtCd()

                'コメントコード（朝）
                .CmtCdM = pfGetCmtCdM(objRecOrderSTD, objRecOMain, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                'コメントコード（10時）
                .CmtCd10 = pfGetCmtCd10(objRecOrderSTD, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                'コメントコード（昼）
                .CmtCdL = pfGetCmtCdL(objRecOrderSTD, objRecOMain, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                'コメントコード（15時）
                .CmtCd15 = pfGetCmtCd15(objRecOrderSTD, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                'コメントコード（夕）
                .CmtCdE = pfGetCmtCdE(objRecOrderSTD, objRecOMain, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                'コメントコード（20時）
                .CmtCd20 = pfGetCmtCd20(objRecOrderSTD, objCmtNorMHash, objCmtKinMHash, blnDelimiterCmtKin, objAllCmtAjiHash, objBumonCmtNorMHash, objBumonCmtKinMHash)

                '共通病名コード
                .IllCd = pfGetIllCd(objRecOrderSTD)

                '病名コード（朝）
                .IllCdM = pfGetIllCdM(objRecOrderSTD)

                '病名コード（10時）
                .IllCd10 = pfGetIllCd10(objRecOrderSTD)

                '病名コード（昼）
                .IllCdL = pfGetIllCdL(objRecOrderSTD)

                '病名コード（15時）
                .IllCd15 = pfGetIllCd15(objRecOrderSTD)

                '病名コード（夕）
                .IllCdE = pfGetIllCdE(objRecOrderSTD)

                '病名コード（20時）
                .IllCd20 = pfGetIllCd20(objRecOrderSTD)

                '栄養成分
                .NtCd = pfGetNtCd(objRecOrderSTD)
                .NtCdM = pfGetNtCdM(objRecOrderSTD)
                .NtCdL = pfGetNtCdL(objRecOrderSTD)
                .NtCdE = pfGetNtCdE(objRecOrderSTD)

                '栄養成分値
                .LmtVal = pfGetLmtVal(objRecOrderSTD)
                .LmtValM = pfGetLmtValM(objRecOrderSTD)
                .LmtValL = pfGetLmtValL(objRecOrderSTD)
                .LmtValE = pfGetLmtValE(objRecOrderSTD)

                ''経管
                .Nut = pfGetNutriTube(objRecOrderSTD, SysInf.OrderCnvTube)

                If .Nut.Trim = "" Then

                    .NutM = pfGetNutriTubeM(objRecOrderSTD, SysInf)
                    .NutL = pfGetNutriTubeL(objRecOrderSTD, SysInf)
                    .NutE = pfGetNutriTubeE(objRecOrderSTD, SysInf)

                Else

                    .NutM = .Nut
                    .NutL = .Nut
                    .NutE = .Nut

                End If

                .Nut10 = ""
                .Nut15 = ""
                .Nut20 = ""

                ''調乳
                .Milk = pfGetMilk(objRecOrderSTD, SysInf.OrderCnvMilk)

                If .Milk.Trim = "" Then

                    .MilkM = pfGetMilkM(objRecOrderSTD, SysInf.OrderCnvMilk)
                    .MilkL = pfGetMilkL(objRecOrderSTD, SysInf.OrderCnvMilk)
                    .MilkE = pfGetMilkE(objRecOrderSTD, SysInf.OrderCnvMilk)

                Else

                    .MilkM = .Milk
                    .MilkL = .Milk
                    .MilkE = .Milk

                End If

                .Milk10 = ""
                .Milk15 = ""
                .Milk20 = ""

                ''静脈
                .VenaM = ""
                .Vena10 = ""
                .VenaL = ""
                .Vena15 = ""
                .VenaE = ""
                .Vena20 = ""

                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                .ErrTp = 0              'エラー区分
                .ErrNo = 0              'エラー番号
                .ErrMsg = ""            'エラーメッセージ

                .UpDtUser = objRecOrderSTD.UpDtUser      '更新者

            End With

            Return objRecOMain

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangeTimeToTmDiv
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
    Public Function ChangeTimeToTmDiv(ByVal strTime As String, _
                                      ByVal SysInf As Rec_XSys) As Integer

        Dim intRtn As Integer

        Try

            Dim strLunch As String
            Dim strDinner As String

            strLunch = SysInf.OrderDlvTmLunch
            strDinner = SysInf.OrderDlvTmDinner

            If strLunch = "" Then
                strLunch = "1000"
            End If

            If strDinner = "" Then
                strDinner = "1600"
            End If

            ''時間が入力されてない場合や不明の名
            If strTime = "" OrElse strTime > "2359" Then

                intRtn = ClsCst.enmEatTimeDiv.MorningDiv

            Else

                If strTime < strLunch Then

                    ''昼食締めの前は朝
                    intRtn = ClsCst.enmEatTimeDiv.MorningDiv

                ElseIf strTime < strDinner Then

                    ''夕食締めの前は昼
                    intRtn = ClsCst.enmEatTimeDiv.LunchDiv

                Else

                    ''夕締めの後は夕
                    intRtn = ClsCst.enmEatTimeDiv.EveningDiv

                End If

            End If

            Return intRtn

        Catch ex As Exception

            Return ClsCst.enmEatTimeDiv.MorningDiv

        End Try

    End Function


    '**********************************************************************
    ' Name       :SplitOrderData
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
    Public Function SplitOrderDataEatTmDiv(ByVal strVal As String, _
                                           ByVal intEatTmDiv As Integer, _
                                           ByVal strSplit As String, _
                                  Optional ByVal blnIsNumeric As Boolean = False) As String

        Dim strRtn As String

        Try

            strRtn = ""

            ''区切り文字がない場合
            If strVal.IndexOf(strSplit) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit, Char))

                If strRtns.Count = 1 Then

                    strRtn = strRtns(0)

                ElseIf strRtns.Count = 3 Then

                    ''朝・昼・夕
                    Select Case intEatTmDiv

                        Case ClsCst.EatTime.MorningDiv, ClsCst.EatTime.AM10Div

                            strRtn = strRtns(0)

                        Case ClsCst.EatTime.LunchDiv, ClsCst.EatTime.PM15Div

                            strRtn = strRtns(1)

                        Case ClsCst.EatTime.EveningDiv, ClsCst.EatTime.NightDiv

                            strRtn = strRtns(2)

                        Case Else

                            strRtn = ""

                    End Select

                ElseIf strRtns.Count = 6 Then

                    ''朝・間・昼・間・夕・間
                    Select Case intEatTmDiv

                        Case ClsCst.EatTime.MorningDiv

                            strRtn = strRtns(0)

                        Case ClsCst.EatTime.AM10Div

                            strRtn = strRtns(1)

                        Case ClsCst.EatTime.LunchDiv

                            strRtn = strRtns(2)

                        Case ClsCst.EatTime.PM15Div

                            strRtn = strRtns(3)

                        Case ClsCst.EatTime.EveningDiv

                            strRtn = strRtns(4)

                        Case ClsCst.EatTime.NightDiv

                            strRtn = strRtns(5)

                        Case Else

                            strRtn = ""

                    End Select

                Else

                    ''不明の場合
                    strRtn = ""

                End If

            Else
                strRtn = strVal
            End If

            ''数値の場合は空白ではなく、0を返す
            If blnIsNumeric = True AndAlso strRtn = "" Then
                strRtn = "0"
            End If

            Return strRtn

        Catch ex As Exception

            strRtn = ""

            If blnIsNumeric = True Then
                strRtn = "0"
            End If

            Return strRtn

        End Try

    End Function

    '**********************************************************************
    ' Name       :SplitOrderData
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
    Public Function SplitOrderData(ByVal strVal As String, _
                                   ByVal strSplit As String) As StringCollection

        Dim objRtnCol As New StringCollection

        Try

            ''区切り文字がない場合
            If strVal.IndexOf(strSplit) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit, Char))

                For Each strWkVal As String In strRtns
                    objRtnCol.Add(strWkVal)
                Next

            Else
                If strVal.Trim <> "" Then
                    objRtnCol.Add(strVal)
                End If
            End If

            Return objRtnCol

        Catch ex As Exception

            Return Nothing

        End Try

    End Function


    '**********************************************************************
    ' Name       :SplitOrderData
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
    Public Function SplitOrderData2Char(ByVal strVal As String, _
                                        ByVal strSplit1 As String, _
                                        ByVal strSplit2 As String) As StringCollection

        Dim objRtnCol As New StringCollection

        Try

            ''区切り文字がない場合
            If strVal.IndexOf(strSplit1) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit1, Char))

                For Each strWkVal As String In strRtns
                    objRtnCol.Add(strWkVal)
                Next

            ElseIf strVal.IndexOf(strSplit2) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit2, Char))

                For Each strWkVal As String In strRtns
                    objRtnCol.Add(strWkVal)
                Next

            Else

                If strVal.Trim <> "" Then
                    objRtnCol.Add(strVal)
                End If
            End If

            Return objRtnCol

        Catch ex As Exception

            Return Nothing

        End Try

    End Function

    '**********************************************************************
    ' Name       :SplitOrderDataGetFirst
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
    Public Function SplitOrderDataGetFirst(ByVal strVal As String, _
                                           ByVal strSplit As String) As String

        Dim objRtnVal As String = ""

        Try

            ''区切り文字がある場合
            If strVal.IndexOf(strSplit) <> -1 Then

                Dim strRtns As String() = strVal.Split(CType(strSplit, Char))

                For Each strWkVal As String In strRtns
                    objRtnVal = strWkVal
                    Exit For
                Next

            Else

                ''区切り文字がない場合
                If strVal.Trim <> "" Then
                    objRtnVal = strVal
                End If
            End If

            Return objRtnVal

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetPreviousTimeDivYMD
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
    Public Function GetPreviousTimeDivYMD(ByVal strYMD As String, _
                                          ByVal intEatTmDiv As Integer) As String

        Dim strNewYMD As String
        Dim dteWorkYMD As Date

        Try

            strNewYMD = ""

            If intEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv Then

                ''朝の場合
                dteWorkYMD = Date.Parse(CType(strYMD, Integer).ToString("0000/00/00"))

                ''前日を取得
                dteWorkYMD = dteWorkYMD.AddDays(-1)

                strNewYMD = dteWorkYMD.ToString("yyyyMMdd")

            Else

                strNewYMD = strYMD

            End If

            Return strNewYMD

        Catch ex As Exception

            Return strYMD

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetPreviousEatTimeDiv
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
    Public Function GetPreviousEatTimeDiv(ByVal intEatTmDiv As Integer, _
                                          ByVal intSnack10 As Integer, _
                                          ByVal intSnack15 As Integer, _
                                          ByVal intSnack20 As Integer) As Integer

        Dim intNewEatTmDiv As Integer

        Try

            If intSnack10 = ClsCst.enmValidF.None AndAlso _
               intSnack15 = ClsCst.enmValidF.None AndAlso _
               intSnack20 = ClsCst.enmValidF.None Then

                ''間食がない場合
                Select Case intEatTmDiv

                    Case ClsCst.enmEatTimeDiv.MorningDiv, ClsCst.enmEatTimeDiv.AM10Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv

                    Case ClsCst.enmEatTimeDiv.LunchDiv, ClsCst.enmEatTimeDiv.PM15Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv

                    Case ClsCst.enmEatTimeDiv.EveningDiv, ClsCst.enmEatTimeDiv.NightDiv

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv

                End Select

            Else

                ''間食がある場合
                Select Case intEatTmDiv

                    Case ClsCst.enmEatTimeDiv.MorningDiv

                        If intSnack20 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.NightDiv
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.AM10Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv

                    Case ClsCst.enmEatTimeDiv.LunchDiv

                        If intSnack10 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.AM10Div
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.PM15Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv

                    Case ClsCst.enmEatTimeDiv.EveningDiv

                        If intSnack15 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.PM15Div
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.NightDiv

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv

                End Select

            End If

            Return intNewEatTmDiv

        Catch ex As Exception

            Return intEatTmDiv

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetNextTimeDivYMD
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
    Public Function GetNextTimeDivYMD(ByVal strYMD As String, _
                                      ByVal intEatTmDiv As Integer, _
                                      ByVal intSnack10 As Integer, _
                                      ByVal intSnack15 As Integer, _
                                      ByVal intSnack20 As Integer) As String

        Dim strNewYMD As String
        Dim dteWorkYMD As Date

        Try

            strNewYMD = ""

            ''20時の間食がない場合
            If (intSnack20 = ClsCst.enmValidF.None AndAlso intEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv) _
               Or (intSnack20 = ClsCst.enmValidF.Exists AndAlso intEatTmDiv = ClsCst.enmEatTimeDiv.NightDiv) Then

                ''夜もしくは夜食の場合
                dteWorkYMD = Date.Parse(CType(strYMD, Integer).ToString("0000/00/00"))

                ''翌日を取得
                dteWorkYMD = dteWorkYMD.AddDays(+1)

                strNewYMD = dteWorkYMD.ToString("yyyyMMdd")

            Else

                strNewYMD = strYMD

            End If

            Return strNewYMD

        Catch ex As Exception

            Return strYMD

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetPreviousEatTimeDiv
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
    Public Function GetNextEatTimeDiv(ByVal intEatTmDiv As Integer, _
                                      ByVal intSnack10 As Integer, _
                                      ByVal intSnack15 As Integer, _
                                      ByVal intSnack20 As Integer) As Integer

        Dim intNewEatTmDiv As Integer

        Try

            If intSnack10 = ClsCst.enmValidF.None AndAlso _
               intSnack15 = ClsCst.enmValidF.None AndAlso _
               intSnack20 = ClsCst.enmValidF.None Then

                ''間食がない場合
                Select Case intEatTmDiv

                    Case ClsCst.enmEatTimeDiv.MorningDiv, ClsCst.enmEatTimeDiv.AM10Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv

                    Case ClsCst.enmEatTimeDiv.LunchDiv, ClsCst.enmEatTimeDiv.PM15Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv

                    Case ClsCst.enmEatTimeDiv.EveningDiv, ClsCst.enmEatTimeDiv.NightDiv

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv

                End Select

            Else

                ''間食がある場合
                Select Case intEatTmDiv

                    Case ClsCst.enmEatTimeDiv.MorningDiv

                        If intSnack10 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.AM10Div
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.AM10Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.LunchDiv

                    Case ClsCst.enmEatTimeDiv.LunchDiv

                        If intSnack15 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.PM15Div
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.PM15Div

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.EveningDiv

                    Case ClsCst.enmEatTimeDiv.EveningDiv

                        If intSnack20 = ClsCst.enmValidF.Exists Then
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.NightDiv
                        Else
                            intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv
                        End If

                    Case ClsCst.enmEatTimeDiv.NightDiv

                        intNewEatTmDiv = ClsCst.enmEatTimeDiv.MorningDiv

                End Select

            End If

            Return intNewEatTmDiv

        Catch ex As Exception

            Return intEatTmDiv

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetNewTelMDiv
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
    Public Function GetNewTelMDiv(ByVal objNowRecOMain As Rec_OMain, _
                                  ByVal objOthRecOMain As Rec_OMain) As String

        Dim objRtn As String = ""

        Try

            If objNowRecOMain.TelMDiv = objOthRecOMain.TelMDiv Then

                objRtn = objNowRecOMain.TelMDiv

            Else

                If objNowRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NS OrElse _
                   objOthRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NS Then

                    objRtn = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NS

                ElseIf objNowRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NI OrElse _
                       objOthRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NI Then

                    objRtn = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_NI

                ElseIf objNowRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_KJ OrElse _
                       objOthRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_KJ Then

                    objRtn = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_KJ

                ElseIf objNowRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_CE OrElse _
                       objOthRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_CE Then

                    objRtn = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_CE

                ElseIf objNowRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_ME OrElse _
                       objOthRecOMain.TelMDiv = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_ME Then

                    objRtn = ClsCst.ORecv.OrderEGMAINGX_TelMDiv_ME

                End If

            End If

            Return objRtn

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :GetNewComOdDiv
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
    Public Function GetNewComOdDiv(ByVal objNowRecOMain As Rec_OMain, _
                                   ByVal objOthRecOMain As Rec_OMain) As String

        Dim objRtn As String = ""

        Try

            If objNowRecOMain.ComOdDiv = objOthRecOMain.ComOdDiv Then

                objRtn = objNowRecOMain.ComOdDiv

            Else

                If objNowRecOMain.ComOdDiv = ClsCst.ORecv.OrderComOdDiv_ComeBack OrElse _
                   objOthRecOMain.ComOdDiv = ClsCst.ORecv.OrderComOdDiv_EatChange Then

                    objRtn = ClsCst.ORecv.OrderComOdDiv_ComeBackEat

                End If

            End If

            Return objRtn

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNutriTube
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
    Private Function pfGetNutriTube(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                    ByVal strAddVal As String) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String

            strRtn = ""

            ''値がない場合もしくは経管の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 2 Then

                ''経管コードがある場合
                If objRecOrderSTD.NutCd.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCd.Count - 1 Step 1

                        strWorkLmtVal = ""

                        ''経管栄養食コード(変換)
                        strWorkLmtVal = strAddVal & pfChangTubeNut(CType(objRecOrderSTD.NutCd.Item(intWkCnt), String))

                        ''給与回数
                        If intWkCnt > objRecOrderSTD.NutMealCnt.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCnt.Item(intWkCnt), String)
                        End If

                        ''給与方法
                        If intWkCnt > objRecOrderSTD.NutMeal.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMeal.Item(intWkCnt), String))
                        End If

                        ''1回量
                        If intWkCnt > objRecOrderSTD.NutOneVal.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneVal.Item(intWkCnt), String)
                        End If

                        ''配膳量
                        If intWkCnt > objRecOrderSTD.NutDlvVal.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvVal.Item(intWkCnt), String)
                        End If

                        ''使用量
                        If intWkCnt > objRecOrderSTD.NutUseVal.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseVal.Item(intWkCnt), String)
                        End If

                        ''備考
                        If intWkCnt > objRecOrderSTD.NutRefCol.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefCol.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkLmtVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNutriTube
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
    Private Function pfGetNutriTubeM(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                     ByVal SysInf As Rec_XSys) As String

        Dim strRtn As String
        Dim strAddVal As String

        Try

            strAddVal = SysInf.OrderCnvTube
            Dim strWorkLmtVal As String

            strRtn = ""

            ''値がない場合もしくは経管の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 2 Then

                ''経管コードがある場合
                If objRecOrderSTD.NutCdM.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdM.Count - 1 Step 1

                        strWorkLmtVal = ""

                        ''経管栄養食コード(変換)
                        strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdM.Item(intWkCnt), String))

                        ''給与回数
                        If intWkCnt > objRecOrderSTD.NutMealCntM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntM.Item(intWkCnt), String)
                        End If

                        ''給与方法
                        If intWkCnt > objRecOrderSTD.NutMealM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealM.Item(intWkCnt), String))
                        End If

                        ''1回量
                        If intWkCnt > objRecOrderSTD.NutOneValM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValM.Item(intWkCnt), String)
                        End If

                        ''配膳量
                        If intWkCnt > objRecOrderSTD.NutDlvValM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValM.Item(intWkCnt), String)
                        End If

                        ''使用量
                        If intWkCnt > objRecOrderSTD.NutUseValM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValM.Item(intWkCnt), String)
                        End If

                        ''備考
                        If intWkCnt > objRecOrderSTD.NutRefColM.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColM.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkLmtVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                        End If

                    Next

                End If

                ''時間別の経管コードがある場合
                If objRecOrderSTD.NutCdTime.Count <> 0 Then

                    Dim strWkTime As String = ""

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdTime.Count - 1 Step 1

                        ''時間を見る
                        strWkTime = (CType(objRecOrderSTD.NutTimeTm.Item(intWkCnt), String))

                        ''朝～昼の間なら
                        If strWkTime <> "" AndAlso strWkTime >= SysInf.OrderDlvTmBreakfast _
                                           AndAlso strWkTime < SysInf.OrderDlvTmLunch Then

                            strWorkLmtVal = ""

                            ''経管栄養食コード(変換)
                            strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdTime.Item(intWkCnt), String))

                            ''給与回数
                            If intWkCnt > objRecOrderSTD.NutMealCntTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntTime.Item(intWkCnt), String)
                            End If

                            ''給与方法
                            If intWkCnt > objRecOrderSTD.NutMealTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealTime.Item(intWkCnt), String))
                            End If

                            ''1回量
                            If intWkCnt > objRecOrderSTD.NutOneValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValTime.Item(intWkCnt), String)
                            End If

                            ''配膳量
                            If intWkCnt > objRecOrderSTD.NutDlvValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValTime.Item(intWkCnt), String)
                            End If

                            ''使用量
                            If intWkCnt > objRecOrderSTD.NutUseValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValTime.Item(intWkCnt), String)
                            End If

                            ''備考
                            If intWkCnt > objRecOrderSTD.NutRefColTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColTime.Item(intWkCnt), String)
                            End If

                            If strRtn = "" Then
                                strRtn = strWorkLmtVal
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                            End If

                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNutriTubeL
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
    Private Function pfGetNutriTubeL(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                     ByVal SysInf As Rec_XSys) As String

        Dim strRtn As String
        Dim strAddVal As String

        Try

            strAddVal = SysInf.OrderCnvTube
            Dim strWorkLmtVal As String

            strRtn = ""

            ''値がない場合もしくは経管の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 2 Then

                ''経管コードがある場合
                If objRecOrderSTD.NutCdL.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdL.Count - 1 Step 1

                        strWorkLmtVal = ""

                        ''経管栄養食コード(変換)
                        strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdL.Item(intWkCnt), String))

                        ''給与回数
                        If intWkCnt > objRecOrderSTD.NutMealCntL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntL.Item(intWkCnt), String)
                        End If

                        ''給与方法
                        If intWkCnt > objRecOrderSTD.NutMealL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealL.Item(intWkCnt), String))
                        End If

                        ''1回量
                        If intWkCnt > objRecOrderSTD.NutOneValL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValL.Item(intWkCnt), String)
                        End If

                        ''配膳量
                        If intWkCnt > objRecOrderSTD.NutDlvValL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValL.Item(intWkCnt), String)
                        End If

                        ''使用量
                        If intWkCnt > objRecOrderSTD.NutUseValL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValL.Item(intWkCnt), String)
                        End If

                        ''備考
                        If intWkCnt > objRecOrderSTD.NutRefColL.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColL.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkLmtVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                        End If

                    Next

                End If

                ''時間別の経管コードがある場合
                If objRecOrderSTD.NutCdTime.Count <> 0 Then

                    Dim strWkTime As String = ""

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdTime.Count - 1 Step 1

                        ''時間を見る
                        strWkTime = (CType(objRecOrderSTD.NutTimeTm.Item(intWkCnt), String))

                        ''昼～夕の間なら
                        If strWkTime <> "" AndAlso strWkTime >= SysInf.OrderDlvTmLunch _
                                           AndAlso strWkTime < SysInf.OrderDlvTmDinner Then

                            strWorkLmtVal = ""

                            ''経管栄養食コード(変換)
                            strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdTime.Item(intWkCnt), String))

                            ''給与回数
                            If intWkCnt > objRecOrderSTD.NutMealCntTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntTime.Item(intWkCnt), String)
                            End If

                            ''給与方法
                            If intWkCnt > objRecOrderSTD.NutMealTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealTime.Item(intWkCnt), String))
                            End If

                            ''1回量
                            If intWkCnt > objRecOrderSTD.NutOneValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValTime.Item(intWkCnt), String)
                            End If

                            ''配膳量
                            If intWkCnt > objRecOrderSTD.NutDlvValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValTime.Item(intWkCnt), String)
                            End If

                            ''使用量
                            If intWkCnt > objRecOrderSTD.NutUseValTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValTime.Item(intWkCnt), String)
                            End If

                            ''備考
                            If intWkCnt > objRecOrderSTD.NutRefColTime.Count - 1 Then
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                            Else
                                strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColTime.Item(intWkCnt), String)
                            End If

                            If strRtn = "" Then
                                strRtn = strWorkLmtVal
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                            End If

                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNutriTubeL
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
    Private Function pfGetNutriTubeE(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                     ByVal SysInf As Rec_XSys) As String

        Dim strRtn As String
        Dim strAddVal As String

        Try

            strAddVal = SysInf.OrderCnvTube
            Dim strWorkLmtVal As String

            strRtn = ""

            ''値がない場合もしくは経管の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 2 Then

                ''経管コードがある場合
                If objRecOrderSTD.NutCdE.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdE.Count - 1 Step 1

                        strWorkLmtVal = ""

                        ''経管栄養食コード（変換）
                        strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdE.Item(intWkCnt), String))

                        ''給与回数
                        If intWkCnt > objRecOrderSTD.NutMealCntE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntE.Item(intWkCnt), String)
                        End If

                        ''給与方法
                        If intWkCnt > objRecOrderSTD.NutMealE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealE.Item(intWkCnt), String))
                        End If

                        ''1回量
                        If intWkCnt > objRecOrderSTD.NutOneValE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValE.Item(intWkCnt), String)
                        End If

                        ''配膳量
                        If intWkCnt > objRecOrderSTD.NutDlvValE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValE.Item(intWkCnt), String)
                        End If

                        ''使用量
                        If intWkCnt > objRecOrderSTD.NutUseValE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValE.Item(intWkCnt), String)
                        End If

                        ''備考
                        If intWkCnt > objRecOrderSTD.NutRefColE.Count - 1 Then
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColE.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkLmtVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                        End If

                    Next

                End If


                ''時間別の経管コードがある場合
                If objRecOrderSTD.NutCdTime.Count <> 0 Then

                    Dim strWkTime As String = ""

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.NutCdTime.Count - 1 Step 1

                        ''時間を見る
                        strWkTime = (CType(objRecOrderSTD.NutTimeTm.Item(intWkCnt), String))

                        ''昼～夕の間なら
                        If strWkTime <> "" Then

                            If strWkTime >= SysInf.OrderDlvTmDinner OrElse _
                               strWkTime < SysInf.OrderDlvTmBreakfast Then

                                strWorkLmtVal = ""

                                ''経管栄養食コード(変換)
                                strWorkLmtVal = strAddVal & Me.pfChangTubeNut(CType(objRecOrderSTD.NutCdTime.Item(intWkCnt), String))

                                ''給与回数
                                If intWkCnt > objRecOrderSTD.NutMealCntTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutMealCntTime.Item(intWkCnt), String)
                                End If

                                ''給与方法
                                If intWkCnt > objRecOrderSTD.NutMealTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & pfChangTubeMealMethod(CType(objRecOrderSTD.NutMealTime.Item(intWkCnt), String))
                                End If

                                ''1回量
                                If intWkCnt > objRecOrderSTD.NutOneValTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutOneValTime.Item(intWkCnt), String)
                                End If

                                ''配膳量
                                If intWkCnt > objRecOrderSTD.NutDlvValTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutDlvValTime.Item(intWkCnt), String)
                                End If

                                ''使用量
                                If intWkCnt > objRecOrderSTD.NutUseValTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutUseValTime.Item(intWkCnt), String)
                                End If

                                ''備考
                                If intWkCnt > objRecOrderSTD.NutRefColTime.Count - 1 Then
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                                Else
                                    strWorkLmtVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NutRefColTime.Item(intWkCnt), String)
                                End If

                                If strRtn = "" Then
                                    strRtn = strWorkLmtVal
                                Else
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkLmtVal
                                End If

                            End If

                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetMilk
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
    Private Function pfGetMilk(ByVal objRecOrderSTD As Rec_OrderSTD, _
                               ByVal strAddVal As String) As String

        Dim strRtn As String

        Try

            Dim strWorkMilkVal As String

            strRtn = ""

            ''値がない場合もしくは調理乳の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 1 Then

                ''調乳コードがある場合
                If objRecOrderSTD.MilkCd.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.MilkCd.Count - 1 Step 1

                        strWorkMilkVal = ""

                        ''調乳コード（変換済）
                        strWorkMilkVal = strAddVal & pfChangMilk(CType(objRecOrderSTD.MilkCd.Item(intWkCnt), String))

                        ''ミルク濃度
                        If intWkCnt > objRecOrderSTD.MilkCon.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkCon.Item(intWkCnt), String)
                        End If

                        ''ミルク量
                        If intWkCnt > objRecOrderSTD.MilkVal.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkVal.Item(intWkCnt), String)
                        End If

                        ''ミルク本数
                        If intWkCnt > objRecOrderSTD.MilkNum.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkNum.Item(intWkCnt), String)
                        End If

                        ''回数
                        If intWkCnt > objRecOrderSTD.MilkMealCnt.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkMealCnt.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶
                        If intWkCnt > objRecOrderSTD.BabyBot.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BabyBot.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶数
                        If intWkCnt > objRecOrderSTD.BotNum.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BotNum.Item(intWkCnt), String)
                        End If

                        ''乳首
                        If intWkCnt > objRecOrderSTD.Nipple.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.Nipple.Item(intWkCnt), String)
                        End If

                        ''乳首数
                        If intWkCnt > objRecOrderSTD.NippNum.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippNum.Item(intWkCnt), String)
                        End If

                        ''希釈方法
                        If intWkCnt > objRecOrderSTD.Dilution.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.Dilution.Item(intWkCnt), String)
                        End If

                        ''ミルク備考
                        If intWkCnt > objRecOrderSTD.MilkRefCol.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkRefCol.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkMilkVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkMilkVal
                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetMilk
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
    Private Function pfGetMilkM(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                ByVal strAddVal As String) As String

        Dim strRtn As String

        Try

            Dim strWorkMilkVal As String

            strRtn = ""

            ''値がない場合もしくは調理乳の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 1 Then

                ''調乳コードがある場合
                If objRecOrderSTD.MilkCdM.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.MilkCdM.Count - 1 Step 1

                        strWorkMilkVal = ""

                        ''調乳コード（変換済）
                        strWorkMilkVal = strAddVal & pfChangMilk(CType(objRecOrderSTD.MilkCdM.Item(intWkCnt), String))

                        ''ミルク濃度
                        If intWkCnt > objRecOrderSTD.MilkConM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkConM.Item(intWkCnt), String)
                        End If

                        ''ミルク量
                        If intWkCnt > objRecOrderSTD.MilkValM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkValM.Item(intWkCnt), String)
                        End If

                        ''ミルク本数
                        If intWkCnt > objRecOrderSTD.MilkNumM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkNumM.Item(intWkCnt), String)
                        End If

                        ''回数
                        If intWkCnt > objRecOrderSTD.MilkMealCntM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkMealCntM.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶
                        If intWkCnt > objRecOrderSTD.BabyBotM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BabyBotM.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶数
                        If intWkCnt > objRecOrderSTD.BotNumM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BotNumM.Item(intWkCnt), String)
                        End If

                        ''乳首
                        If intWkCnt > objRecOrderSTD.NippleM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippleM.Item(intWkCnt), String)
                        End If

                        ''乳首数
                        If intWkCnt > objRecOrderSTD.NippNumM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippNumM.Item(intWkCnt), String)
                        End If

                        ''希釈方法
                        If intWkCnt > objRecOrderSTD.DilutionM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.DilutionM.Item(intWkCnt), String)
                        End If

                        ''ミルク備考
                        If intWkCnt > objRecOrderSTD.MilkRefColM.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkRefColM.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkMilkVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkMilkVal
                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetMilk
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
    Private Function pfGetMilkL(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                ByVal strAddVal As String) As String

        Dim strRtn As String

        Try

            Dim strWorkMilkVal As String

            strRtn = ""

            ''値がない場合もしくは調理乳の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 1 Then

                ''調乳コードがある場合
                If objRecOrderSTD.MilkCdL.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.MilkCdL.Count - 1 Step 1

                        strWorkMilkVal = ""

                        ''調乳コード（変換済）
                        strWorkMilkVal = strAddVal & pfChangMilk(CType(objRecOrderSTD.MilkCdL.Item(intWkCnt), String))

                        ''ミルク濃度
                        If intWkCnt > objRecOrderSTD.MilkConL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkConL.Item(intWkCnt), String)
                        End If

                        ''ミルク量
                        If intWkCnt > objRecOrderSTD.MilkValL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkValL.Item(intWkCnt), String)
                        End If

                        ''ミルク本数
                        If intWkCnt > objRecOrderSTD.MilkNumL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkNumL.Item(intWkCnt), String)
                        End If

                        ''回数
                        If intWkCnt > objRecOrderSTD.MilkMealCntL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkMealCntL.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶
                        If intWkCnt > objRecOrderSTD.BabyBotL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BabyBotL.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶数
                        If intWkCnt > objRecOrderSTD.BotNumL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BotNumL.Item(intWkCnt), String)
                        End If

                        ''乳首
                        If intWkCnt > objRecOrderSTD.NippleL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippleL.Item(intWkCnt), String)
                        End If

                        ''乳首数
                        If intWkCnt > objRecOrderSTD.NippNumL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippNumL.Item(intWkCnt), String)
                        End If

                        ''希釈方法
                        If intWkCnt > objRecOrderSTD.DilutionL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.DilutionL.Item(intWkCnt), String)
                        End If

                        ''ミルク備考
                        If intWkCnt > objRecOrderSTD.MilkRefColL.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkRefColL.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkMilkVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkMilkVal
                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfGetMilk
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
    Private Function pfGetMilkE(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                ByVal strAddVal As String) As String

        Dim strRtn As String

        Try

            Dim strWorkMilkVal As String

            strRtn = ""

            ''値がない場合もしくは調理乳の場合
            If objRecOrderSTD.HaveMilkTube = -1 _
               OrElse objRecOrderSTD.HaveMilkTube = 0 _
               OrElse objRecOrderSTD.HaveMilkTube = 1 Then

                ''調乳コードがある場合
                If objRecOrderSTD.MilkCdE.Count <> 0 Then

                    ''
                    For intWkCnt As Integer = 0 To objRecOrderSTD.MilkCdE.Count - 1 Step 1

                        strWorkMilkVal = ""

                        ''調乳コード（変換済）
                        strWorkMilkVal = strAddVal & pfChangMilk(CType(objRecOrderSTD.MilkCdE.Item(intWkCnt), String))

                        ''ミルク濃度
                        If intWkCnt > objRecOrderSTD.MilkConE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkConE.Item(intWkCnt), String)
                        End If

                        ''ミルク量
                        If intWkCnt > objRecOrderSTD.MilkValE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkValE.Item(intWkCnt), String)
                        End If

                        ''ミルク本数
                        If intWkCnt > objRecOrderSTD.MilkNumE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkNumE.Item(intWkCnt), String)
                        End If

                        ''回数
                        If intWkCnt > objRecOrderSTD.MilkMealCntE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkMealCntE.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶
                        If intWkCnt > objRecOrderSTD.BabyBotE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BabyBotE.Item(intWkCnt), String)
                        End If

                        ''哺乳瓶数
                        If intWkCnt > objRecOrderSTD.BotNumE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.BotNumE.Item(intWkCnt), String)
                        End If

                        ''乳首
                        If intWkCnt > objRecOrderSTD.NippleE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippleE.Item(intWkCnt), String)
                        End If

                        ''乳首数
                        If intWkCnt > objRecOrderSTD.NippNumE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.NippNumE.Item(intWkCnt), String)
                        End If

                        ''希釈方法
                        If intWkCnt > objRecOrderSTD.DilutionE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.DilutionE.Item(intWkCnt), String)
                        End If

                        ''ミルク備考
                        If intWkCnt > objRecOrderSTD.MilkRefColE.Count - 1 Then
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & ""
                        Else
                            strWorkMilkVal &= ORDER_DELIMITER_CHAR_SEMICOLON & CType(objRecOrderSTD.MilkRefColE.Item(intWkCnt), String)
                        End If

                        If strRtn = "" Then
                            strRtn = strWorkMilkVal
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWorkMilkVal
                        End If

                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfGetLmtVal
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
    Private Function pfGetLmtVal(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.LmtVal.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.LmtVal.Count - 1 Step 1

                    strWorkLmtVal = CType(objRecOrderSTD.LmtVal.Item(intWkCnt), String)

                    If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                        dblWorkLmtVal = 0
                    Else
                        dblWorkLmtVal = CType(strWorkLmtVal, Double)
                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = CType(dblWorkLmtVal, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(dblWorkLmtVal, String)
                        End If

                    End If

                Next

                'For Each strLmtVal As String In objRecOrderSTD.LmtVal

                '    If strRtn = "" Then
                '        strRtn = strLmtVal.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strLmtVal.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetLmtVal
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
    Private Function pfGetLmtValM(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.LmtValM.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.LmtValM.Count - 1 Step 1

                    strWorkLmtVal = CType(objRecOrderSTD.LmtValM.Item(intWkCnt), String)

                    If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                        dblWorkLmtVal = 0
                    Else
                        dblWorkLmtVal = CType(strWorkLmtVal, Double)
                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = CType(dblWorkLmtVal, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(dblWorkLmtVal, String)
                        End If

                    End If

                Next

                'For Each strLmtVal As String In objRecOrderSTD.LmtValM

                '    If strRtn = "" Then
                '        strRtn = strLmtVal.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strLmtVal.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetLmtVal
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
    Private Function pfGetLmtValL(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.LmtValL.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.LmtValL.Count - 1 Step 1

                    strWorkLmtVal = CType(objRecOrderSTD.LmtValL.Item(intWkCnt), String)

                    If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                        dblWorkLmtVal = 0
                    Else
                        dblWorkLmtVal = CType(strWorkLmtVal, Double)
                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = CType(dblWorkLmtVal, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(dblWorkLmtVal, String)
                        End If

                    End If

                Next

                'For Each strLmtVal As String In objRecOrderSTD.LmtValL

                '    If strRtn = "" Then
                '        strRtn = strLmtVal.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strLmtVal.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetLmtVal
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
    Private Function pfGetLmtValE(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.LmtValE.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.LmtValE.Count - 1 Step 1

                    strWorkLmtVal = CType(objRecOrderSTD.LmtValE.Item(intWkCnt), String)

                    If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                        dblWorkLmtVal = 0
                    Else
                        dblWorkLmtVal = CType(strWorkLmtVal, Double)
                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = CType(dblWorkLmtVal, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(dblWorkLmtVal, String)
                        End If

                    End If

                Next

                'For Each strLmtVal As String In objRecOrderSTD.LmtValE

                '    If strRtn = "" Then
                '        strRtn = strLmtVal.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strLmtVal.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNtCd
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
    Private Function pfGetNtCd(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.NtCd.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.NtCd.Count - 1 Step 1

                    ''制限値を確認する
                    If intWkCnt > objRecOrderSTD.LmtVal.Count - 1 Then
                        dblWorkLmtVal = 0
                    Else
                        strWorkLmtVal = CType(objRecOrderSTD.LmtVal.Item(intWkCnt), String)

                        If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                            dblWorkLmtVal = 0
                        Else
                            dblWorkLmtVal = CType(strWorkLmtVal, Double)
                        End If

                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = objRecOrderSTD.NtCd.Item(intWkCnt).Trim
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecOrderSTD.NtCd.Item(intWkCnt).Trim
                        End If

                    End If

                Next

                'For Each strNtCd As String In objRecOrderSTD.NtCd

                '    If strRtn = "" Then
                '        strRtn = strNtCd.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNtCd.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNtCd
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
    Private Function pfGetNtCdM(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try
            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.NtCdM.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.NtCdM.Count - 1 Step 1

                    ''制限値を確認する
                    If intWkCnt > objRecOrderSTD.LmtValM.Count - 1 Then
                        dblWorkLmtVal = 0
                    Else
                        strWorkLmtVal = CType(objRecOrderSTD.LmtValM.Item(intWkCnt), String)

                        If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                            dblWorkLmtVal = 0
                        Else
                            dblWorkLmtVal = CType(strWorkLmtVal, Double)
                        End If

                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = objRecOrderSTD.NtCdM.Item(intWkCnt).Trim
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecOrderSTD.NtCdM.Item(intWkCnt).Trim
                        End If

                    End If

                Next

                'For Each strNtCd As String In objRecOrderSTD.NtCdM

                '    If strRtn = "" Then
                '        strRtn = strNtCd.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNtCd.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNtCd
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
    Private Function pfGetNtCdL(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try
            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.NtCdL.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.NtCdL.Count - 1 Step 1

                    ''制限値を確認する
                    If intWkCnt > objRecOrderSTD.LmtValL.Count - 1 Then
                        dblWorkLmtVal = 0
                    Else
                        strWorkLmtVal = CType(objRecOrderSTD.LmtValL.Item(intWkCnt), String)

                        If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                            dblWorkLmtVal = 0
                        Else
                            dblWorkLmtVal = CType(strWorkLmtVal, Double)
                        End If

                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = objRecOrderSTD.NtCdL.Item(intWkCnt).Trim
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecOrderSTD.NtCdL.Item(intWkCnt).Trim
                        End If

                    End If

                Next

                'For Each strNtCd As String In objRecOrderSTD.NtCdL

                '    If strRtn = "" Then
                '        strRtn = strNtCd.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNtCd.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNtCd
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
    Private Function pfGetNtCdE(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            Dim strWorkLmtVal As String
            Dim dblWorkLmtVal As Double

            strRtn = ""

            If objRecOrderSTD.NtCdE.Count <> 0 Then

                For intWkCnt As Integer = 0 To objRecOrderSTD.NtCdE.Count - 1 Step 1

                    ''制限値を確認する
                    If intWkCnt > objRecOrderSTD.LmtValE.Count - 1 Then
                        dblWorkLmtVal = 0
                    Else
                        strWorkLmtVal = CType(objRecOrderSTD.LmtValE.Item(intWkCnt), String)

                        If strWorkLmtVal = "" OrElse IsNumeric(strWorkLmtVal) = False Then
                            dblWorkLmtVal = 0
                        Else
                            dblWorkLmtVal = CType(strWorkLmtVal, Double)
                        End If

                    End If

                    If dblWorkLmtVal <> 0 Then

                        If strRtn = "" Then
                            strRtn = objRecOrderSTD.NtCdE.Item(intWkCnt).Trim
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecOrderSTD.NtCdE.Item(intWkCnt).Trim
                        End If

                    End If

                Next

                'For Each strNtCd As String In objRecOrderSTD.NtCdE

                '    If strRtn = "" Then
                '        strRtn = strNtCd.Trim
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNtCd.Trim
                '    End If

                'Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtCd20
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
    Private Function pfGetCmtCd20(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                  ByVal objCmtNorMHash As Hashtable, _
                                  ByVal objCmtKinMHash As Hashtable, _
                                  ByVal blnDelimiterKin As Boolean, _
                                  ByVal objAllCmtAjiHash As Hashtable, _
                                  ByVal objBumonCmtNorMHash As Hashtable, _
                                  ByVal objBumonCmtKinMHash As Hashtable) As String


        Dim strRtn As String

        Try

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            Dim strNewCmtCode As String
            Dim blnBumon As Boolean = False

            strRtn = ""

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''禁止コメント（20）
                If objRecOrderSTD.CmtKinCd20.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd20

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（夕）
                If objRecOrderSTD.CmtCd20.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd20

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（20時）
                If objRecOrderSTD.CmtCd20.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd20

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.NightDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtCdE
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
    Private Function pfGetCmtCdE(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objRecOMain As Rec_OMain, _
                                 ByVal objCmtNorMHash As Hashtable, _
                                 ByVal objCmtKinMHash As Hashtable, _
                                 ByVal blnDelimiterKin As Boolean, _
                                 ByVal objAllCmtAjiHash As Hashtable, _
                                 ByVal objBumonCmtNorMHash As Hashtable, _
                                 ByVal objBumonCmtKinMHash As Hashtable) As String

        Dim strRtn As String

        Try

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            Dim strNewCmtCode As String
            Dim blnBumon As Boolean = False

            strRtn = ""

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If


                    Next

                End If

                ''禁止コメント（夕）
                If objRecOrderSTD.CmtKinCdE.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCdE

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If


                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（夕）
                If objRecOrderSTD.CmtCdE.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdE

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（夕）
                If objRecOrderSTD.CmtCdE.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdE

                        ''新しい通常コードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.EveningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''追加のコメントを格納する
            Dim objCmtAdd As New StringCollection

            If objRecOMain.CmtAddedENew = True Then

                If strRtn = "" Then
                    strRtn = objRecOMain.CmtAddENew
                Else

                    ''コメントを追加する（既存のコメントは追加しない）
                    objCmtAdd = SplitOrderData(objRecOMain.CmtAddENew, ORDER_DELIMITER_CHAR_PIPE)

                    For Each strCmt As String In objCmtAdd
                        If strRtn.IndexOf(strCmt) = -1 Then
                            If strRtn = "" Then
                                strRtn = strCmt
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                            End If
                        End If
                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtCd15
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
    Private Function pfGetCmtCd15(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                  ByVal objCmtNorMHash As Hashtable, _
                                  ByVal objCmtKinMHash As Hashtable, _
                                  ByVal blnDelimiterKin As Boolean, _
                                  ByVal objAllCmtAjiHash As Hashtable, _
                                  ByVal objBumonCmtNorMHash As Hashtable, _
                                  ByVal objBumonCmtKinMHash As Hashtable) As String

        Dim strRtn As String

        Try

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            Dim strNewCmtCode As String
            Dim blnBumon As Boolean = False

            strRtn = ""

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''禁止コメント（15時）
                If objRecOrderSTD.CmtKinCd15.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd15

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（15時）
                If objRecOrderSTD.CmtCd15.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd15

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd


                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（15時）
                If objRecOrderSTD.CmtCd15.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd15


                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.PM15Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／
            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtCdL
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
    Private Function pfGetCmtCdL(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objRecOMain As Rec_OMain, _
                                 ByVal objCmtNorMHash As Hashtable, _
                                 ByVal objCmtKinMHash As Hashtable, _
                                 ByVal blnDelimiterKin As Boolean, _
                                 ByVal objAllCmtAjiHash As Hashtable, _
                                 ByVal objBumonCmtNorMHash As Hashtable, _
                                 ByVal objBumonCmtKinMHash As Hashtable) As String

        Dim strRtn As String

        Try

            Dim strNewCmtCode As String
            Dim blnBumon As Boolean = False

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            strRtn = ""

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''禁止コメント（昼）
                If objRecOrderSTD.CmtKinCdL.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCdL

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（昼）
                If objRecOrderSTD.CmtCdL.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdL

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（昼）
                If objRecOrderSTD.CmtCdL.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdL

                        ''新しい通常コードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.LunchDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''追加のコメントを格納する
            Dim objCmtAdd As New StringCollection

            If objRecOMain.CmtAddedLNew = True Then

                If strRtn = "" Then
                    strRtn = objRecOMain.CmtAddLNew
                Else

                    ''コメントを追加する（既存のコメントは追加しない）
                    objCmtAdd = SplitOrderData(objRecOMain.CmtAddLNew, ORDER_DELIMITER_CHAR_PIPE)

                    For Each strCmt As String In objCmtAdd
                        If strRtn.IndexOf(strCmt) = -1 Then
                            If strRtn = "" Then
                                strRtn = strCmt
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                            End If
                        End If
                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetCmtCd10
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
    Private Function pfGetCmtCd10(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                  ByVal objCmtNorMHash As Hashtable, _
                                  ByVal objCmtKinMHash As Hashtable, _
                                  ByVal blnDelimiterKin As Boolean, _
                                  ByVal objAllCmtAjiHash As Hashtable, _
                                  ByVal objBumonCmtNorMHash As Hashtable, _
                                  ByVal objBumonCmtKinMHash As Hashtable) As String

        Dim strRtn As String

        Try

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            Dim strNewCmtCode As String
            Dim blnBumon As Boolean = False

            strRtn = ""

            If blnDelimiterKin = True Then

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''禁止コメント（10時）
                If objRecOrderSTD.CmtKinCd10.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd10

                        ''新しい禁止コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（10時）
                If objRecOrderSTD.CmtCd10.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd10

                        ''新しい通常コメントコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しい通常コードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（10時）
                If objRecOrderSTD.CmtCd10.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd10

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.AM10Div, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを先に格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    ''**********************************************************************
    '' Name       :pfGetCmtCdM
    ''
    ''----------------------------------------------------------------------
    '' Comment    :
    ''
    ''----------------------------------------------------------------------
    '' Design On  :2011/11/08
    ''        By  :FTX)S.Shin
    ''----------------------------------------------------------------------
    '' Update On  :
    ''        By  :
    '' Comment    :
    ''**********************************************************************
    'Private Function pfGetCmtCdM(ByVal objRecOrderSTD As Rec_OrderSTD, _
    '                             ByVal objCmtNorMHash As Hashtable, _
    '                             ByVal objCmtKinMHash As Hashtable) As String

    '    Dim strRtn As String

    '    Try

    '        Dim objRecMCmt As New Rec_MCmt

    '        strRtn = ""

    '        ''禁止コメント（朝）
    '        If objRecOrderSTD.CmtKinCdM.Count <> 0 Then

    '            For Each strCmt As String In objRecOrderSTD.CmtKinCdM

    '                ''マスタにあった場合
    '                If objCmtKinMHash.ContainsKey(strCmt) Then

    '                    objRecMCmt = CType(objCmtKinMHash.Item(strCmt), Rec_MCmt)

    '                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then

    '                        If strRtn = "" Then
    '                            strRtn = objRecMCmt.CmtCd.Trim
    '                        Else
    '                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecMCmt.CmtCd.Trim
    '                        End If

    '                    End If

    '                Else

    '                    ''不明なコメント

    '                End If

    '            Next

    '        End If

    '        ''コメント（朝）
    '        If objRecOrderSTD.CmtCdM.Count <> 0 Then

    '            For Each strCmt As String In objRecOrderSTD.CmtCdM

    '                ''マスタにあった場合
    '                If objCmtNorMHash.ContainsKey(strCmt) Then

    '                    objRecMCmt = CType(objCmtNorMHash.Item(strCmt), Rec_MCmt)

    '                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then

    '                        If strRtn = "" Then
    '                            strRtn = objRecMCmt.CmtCd.Trim
    '                        Else
    '                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecMCmt.CmtCd.Trim
    '                        End If

    '                    End If

    '                Else

    '                    ''不明なコメント

    '                End If

    '            Next

    '        End If

    '        Return strRtn

    '    Catch ex As Exception

    '        Throw

    '    End Try

    'End Function

    '**********************************************************************
    ' Name       :pfGetCmtCdM
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
    Private Function pfGetCmtCdM(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objRecOMain As Rec_OMain, _
                                 ByVal objCmtNorMHash As Hashtable, _
                                 ByVal objCmtKinMHash As Hashtable, _
                                 ByVal blnDelimiterKin As Boolean, _
                                 ByVal objAllCmtAjiHash As Hashtable, _
                                 ByVal objBumonCmtNorMHash As Hashtable, _
                                 ByVal objBumonCmtKinMHash As Hashtable) As String

        Dim strRtn As String

        Try

            Dim strNewCmtCode As String

            Dim objOrderCmt As New StringCollection
            Dim objBumonCmt As New StringCollection

            Dim blnBumon As Boolean = False

            strRtn = ""

            If blnDelimiterKin = True Then
                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを分割して送ってくる場合 　　　　　　　　　＼
                ''
                ''共通禁止コメント
                If objRecOrderSTD.CmtKinCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCd

                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''禁止コメント（朝）
                If objRecOrderSTD.CmtKinCdM.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtKinCdM

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objBumonCmtKinMHash, _
                                                      enumCmtCls.KinCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（朝）
                If objRecOrderSTD.CmtCdM.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdM

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtNorMHash, _
                                                      enumCmtCls.NorCmt, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／
            Else

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／'禁止とコメントを一緒に送ってくる場合   　　　　　　　　　＼
                ''
                ''共通コメント
                If objRecOrderSTD.CmtCd.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCd

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If

                ''コメント（朝）
                If objRecOrderSTD.CmtCdM.Count <> 0 Then

                    For Each strCmt As String In objRecOrderSTD.CmtCdM

                        ''新しいコードを取得
                        strNewCmtCode = pfGetNewCmtCd(strCmt, _
                                                      ClsCst.enmEatTimeDiv.MorningDiv, _
                                                      objAllCmtAjiHash, _
                                                      objCmtKinMHash, _
                                                      objCmtNorMHash, _
                                                      objBumonCmtKinMHash, _
                                                      objBumonCmtNorMHash, blnBumon)

                        If strNewCmtCode <> "" Then

                            If blnBumon = False Then
                                ''オーダコメントの場合
                                objOrderCmt.Add(strNewCmtCode)
                            Else
                                ''部門コメントの場合
                                objBumonCmt.Add(strNewCmtCode)
                            End If

                        End If

                    Next

                End If
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            End If

            ''オーダコメントを先に格納
            For Each strCmt As String In objOrderCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''部門コメントを後に格納
            For Each strCmt As String In objBumonCmt

                If strRtn = "" Then
                    strRtn = strCmt
                Else
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                End If

            Next

            ''追加のコメントを格納する
            Dim objCmtAdd As New StringCollection

            If objRecOMain.CmtAddedMNew = True Then

                If strRtn = "" Then
                    strRtn = objRecOMain.CmtAddMNew
                Else

                    ''コメントを追加する（既存のコメントは追加しない）
                    objCmtAdd = SplitOrderData(objRecOMain.CmtAddMNew, ORDER_DELIMITER_CHAR_PIPE)

                    For Each strCmt As String In objCmtAdd
                        If strRtn.IndexOf(strCmt) = -1 Then
                            If strRtn = "" Then
                                strRtn = strCmt
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
                            End If
                        End If
                    Next

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfCheckAllergyCmt
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
    Private Function pfCheckAllergyCmtDtl(ByVal strAllergyCmt As String, _
                                          ByVal objOrderCmt As StringCollection) As Boolean

        Dim blnRtn As Boolean = False

        Try

            If Not objOrderCmt Is Nothing AndAlso objOrderCmt.Count > 0 Then

                For Each strCmt As String In objOrderCmt

                    If strCmt.Trim = strAllergyCmt.Trim Then
                        blnRtn = True
                        Exit For
                    End If

                Next

            End If

            Return blnRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfCheckAllergyCmt
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
    Private Function pfCheckAllergyCmt(ByVal SysInf As Rec_XSys, _
                                       ByVal objRecOrderSTD As Rec_OrderSTD) As Boolean

        Dim blnRtn As Boolean = False

        Try

            ''アレルギーコメントをチェック
            Dim objAllergyCmt As New StringCollection

            objAllergyCmt = SplitOrderData(SysInf.OrderAllergyCmt, ORDER_DELIMITER_CHAR_PIPE)

            If Not objAllergyCmt Is Nothing AndAlso objAllergyCmt.Count > 0 Then

                For Each strAllergyCmt As String In objAllergyCmt

                    '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                    '／禁止コメントをチェック　　　　　　　　　　　　　　　　　　＼
                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCd)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCdM)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCd10)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCdL)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCd15)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCdE)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtKinCd20)
                    End If
                    '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                    '／コメントをチェック　　　　　　　　　　　　　　　　　　　　＼
                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCd)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCd10)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCdL)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCd15)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCdE)
                    End If

                    If blnRtn = False Then
                        blnRtn = pfCheckAllergyCmtDtl(strAllergyCmt, objRecOrderSTD.CmtCd20)
                    End If
                    '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    If blnRtn = True Then
                        Exit For
                    End If

                Next

            End If

            Return blnRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetDelivCd
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
    Private Function pfGetCmtCd() As String

        Dim strRtn As String

        Try

            strRtn = ""

            'Dim objRecMCmt As New Rec_MCmt

            'Dim objOrderCmt As New StringCollection
            'Dim objBumonCmt As New StringCollection

            'strRtn = ""

            ' ''禁止コメント
            'If objRecOrderSTD.CmtKinCd.Count <> 0 Then

            '    For Each strCmt As String In objRecOrderSTD.CmtKinCd

            '        ''オーダのマスタにあった場合
            '        If objCmtKinMHash.ContainsKey(strCmt) Then

            '            objRecMCmt = CType(objCmtKinMHash.Item(strCmt), Rec_MCmt)

            '            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
            '                objOrderCmt.Add(objRecMCmt.CmtCd.Trim)
            '            End If

            '        Else

            '            ''部門コメント
            '            If Not objBumonCmtKinMHash Is Nothing Then

            '                ''部門マスタにあった場合
            '                If objBumonCmtKinMHash.ContainsKey(strCmt) Then

            '                    objRecMCmt = CType(objBumonCmtKinMHash.Item(strCmt), Rec_MCmt)

            '                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
            '                        objBumonCmt.Add(objRecMCmt.CmtCd.Trim)
            '                    End If

            '                End If

            '            End If

            '        End If

            '    Next

            'End If

            ' ''コメント（朝）
            'If objRecOrderSTD.CmtCdM.Count <> 0 Then

            '    For Each strCmt As String In objRecOrderSTD.CmtCdM

            '        ''オーダのマスタにあった場合
            '        If objCmtNorMHash.ContainsKey(strCmt) Then

            '            objRecMCmt = CType(objCmtNorMHash.Item(strCmt), Rec_MCmt)

            '            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
            '                objOrderCmt.Add(objRecMCmt.CmtCd.Trim)
            '            End If

            '        Else

            '            ''部門コメント
            '            If Not objBumonCmtNorMHash Is Nothing Then

            '                ''部門マスタにあった場合
            '                If objBumonCmtNorMHash.ContainsKey(strCmt) Then

            '                    objRecMCmt = CType(objBumonCmtNorMHash.Item(strCmt), Rec_MCmt)

            '                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
            '                        objBumonCmt.Add(objRecMCmt.CmtCd.Trim)
            '                    End If

            '                End If

            '            End If

            '        End If

            '    Next

            'End If


            ' ''オーダコメントを先に格納
            'For Each strCmt As String In objOrderCmt

            '    If strRtn = "" Then
            '        strRtn = strCmt
            '    Else
            '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
            '    End If

            'Next

            ' ''部門コメントを先に格納
            'For Each strCmt As String In objBumonCmt

            '    If strRtn = "" Then
            '        strRtn = strCmt
            '    Else
            '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strCmt
            '    End If

            'Next

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCd(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCd.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCd

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCdM(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCdM.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCdM

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCd10(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCd10.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCd10

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCdL(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCdL.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCdL

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCd15(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCd15.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCd15

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCdE(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCdE.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCdE

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetIllCd
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
    Private Function pfGetIllCd20(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strIllCd As String

            If objRecOrderSTD.IllCd20.Count <> 0 Then

                For Each strIll As String In objRecOrderSTD.IllCd20

                    ''コードを取得と変換
                    strIllCd = pfChangIllCd(strIll)

                    If strRtn = "" Then
                        strRtn = strIllCd.Trim
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strIllCd.Trim
                    End If

                Next

            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetSelEtDiv
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
    Private Function pfGetSelEtDiv(ByVal objRecOrderSTD As Rec_OrderSTD) As String

        Dim strRtn As String

        Try

            strRtn = ""

            strRtn = objRecOrderSTD.SelEtDiv.Trim

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetDelivCd
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
    Private Function pfGetDelivCd(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                  ByVal blnSnack10 As Boolean, _
                                  ByVal blnSnack15 As Boolean, _
                                  ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strDelivCd As String
            Dim strDelivCdM As String
            Dim strDelivCd10 As String
            Dim strDelivCdL As String
            Dim strDelivCd15 As String
            Dim strDelivCdE As String
            Dim strDelivCd20 As String

            ''コードを取得と変換
            strDelivCd = pfChangDelivCd(objRecOrderSTD.DelivCd.Trim)
            strDelivCdM = pfChangDelivCd(objRecOrderSTD.DelivCdM.Trim)
            strDelivCd10 = pfChangDelivCd(objRecOrderSTD.DelivCd10.Trim)
            strDelivCdL = pfChangDelivCd(objRecOrderSTD.DelivCdL.Trim)
            strDelivCd15 = pfChangDelivCd(objRecOrderSTD.DelivCd15.Trim)
            strDelivCdE = pfChangDelivCd(objRecOrderSTD.DelivCdE.Trim)
            strDelivCd20 = pfChangDelivCd(objRecOrderSTD.DelivCd20.Trim)

            '1日の配膳の場合
            If strDelivCd.Trim <> "" Then

                strRtn = strDelivCd.Trim

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''全てが空白なら
                    If strDelivCdM.Trim = "" AndAlso _
                       strDelivCd10.Trim = "" AndAlso _
                       strDelivCdL.Trim = "" AndAlso _
                       strDelivCd15.Trim = "" AndAlso _
                       strDelivCdE.Trim = "" AndAlso _
                       strDelivCd20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strDelivCdM.Trim

                        '10
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCd10.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCdL.Trim

                        '15
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCd15.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCdE.Trim

                        '20
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCd20.Trim

                    End If
                Else

                    ''全てが空白なら
                    If strDelivCdM.Trim = "" AndAlso _
                       strDelivCdL.Trim = "" AndAlso _
                       strDelivCdE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strDelivCdM.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCdL.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strDelivCdE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetWardCd
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
    Private Function pfGetWardCd(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strWard As String
            Dim strWardM As String
            Dim strWard10 As String
            Dim strWardL As String
            Dim strWard15 As String
            Dim strWardE As String
            Dim strWard20 As String

            ''コードを取得と変換
            strWard = pfChangWard(objRecOrderSTD.Ward.Trim)
            strWardM = pfChangWard(objRecOrderSTD.WardM.Trim)
            strWard10 = pfChangWard(objRecOrderSTD.Ward10.Trim)
            strWardL = pfChangWard(objRecOrderSTD.WardL.Trim)
            strWard15 = pfChangWard(objRecOrderSTD.Ward15.Trim)
            strWardE = pfChangWard(objRecOrderSTD.WardE.Trim)
            strWard20 = pfChangWard(objRecOrderSTD.Ward20.Trim)

            '1日の配膳の場合
            If strWard.Trim <> "" Then

                strRtn = strWard.Trim

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''全てが空白なら
                    If strWardM.Trim = "" AndAlso _
                       strWard10.Trim = "" AndAlso _
                       strWardL.Trim = "" AndAlso _
                       strWard15.Trim = "" AndAlso _
                       strWardE.Trim = "" AndAlso _
                       strWard20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strWardM.Trim

                        '10
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWard10.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWardL.Trim

                        '15
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWard15.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWardE.Trim

                        '20
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWard20.Trim

                    End If

                Else

                    ''全てが空白なら
                    If strWardM.Trim = "" AndAlso _
                       strWardL.Trim = "" AndAlso _
                       strWardE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strWardM.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWardL.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strWardE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetRoodCd
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
    Private Function pfGetRoodCd(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strRoom As String
            Dim strRoomM As String
            Dim strRoom10 As String
            Dim strRoomL As String
            Dim strRoom15 As String
            Dim strRoomE As String
            Dim strRoom20 As String

            ''コードを取得と変換
            strRoom = pfChangRoom(objRecOrderSTD.Room.Trim)
            strRoomM = pfChangRoom(objRecOrderSTD.RoomM.Trim)
            strRoom10 = pfChangRoom(objRecOrderSTD.Room10.Trim)
            strRoomL = pfChangRoom(objRecOrderSTD.RoomL.Trim)
            strRoom15 = pfChangRoom(objRecOrderSTD.Room15.Trim)
            strRoomE = pfChangRoom(objRecOrderSTD.RoomE.Trim)
            strRoom20 = pfChangRoom(objRecOrderSTD.Room20.Trim)

            '1日の配膳の場合
            If strRoom.Trim <> "" Then

                strRtn = strRoom.Trim

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''全てが空白なら
                    If strRoomM.Trim = "" AndAlso _
                       strRoom10.Trim = "" AndAlso _
                       strRoomL.Trim = "" AndAlso _
                       strRoom15.Trim = "" AndAlso _
                       strRoomE.Trim = "" AndAlso _
                       strRoom20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strRoomM.Trim

                        '10
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoom10.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoomL.Trim

                        '15
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoom15.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoomE.Trim

                        '20
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoom20.Trim

                    End If

                Else

                    ''全てが空白なら
                    If strRoomM.Trim = "" AndAlso _
                       strRoomL.Trim = "" AndAlso _
                       strRoomE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strRoomM.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoomL.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strRoomE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfGetRoodCd
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
    Private Function pfGetBed(ByVal objRecOrderSTD As Rec_OrderSTD, _
                              ByVal blnSnack10 As Boolean, _
                              ByVal blnSnack15 As Boolean, _
                              ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String

        Try

            strRtn = ""

            Dim strBed As String
            Dim strBedM As String
            Dim strBed10 As String
            Dim strBedL As String
            Dim strBed15 As String
            Dim strBedE As String
            Dim strBed20 As String

            strBed = objRecOrderSTD.Bed.Trim
            strBedM = objRecOrderSTD.BedM.Trim
            strBed10 = objRecOrderSTD.Bed10.Trim
            strBedL = objRecOrderSTD.BedL.Trim
            strBed15 = objRecOrderSTD.Bed15.Trim
            strBedE = objRecOrderSTD.BedE.Trim
            strBed20 = objRecOrderSTD.Bed20.Trim

            '1日の配膳の場合
            If strBed.Trim <> "" Then

                strRtn = strBed.Trim

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''全てが空白なら
                    If strBedM.Trim = "" AndAlso _
                       strBed10.Trim = "" AndAlso _
                       strBedL.Trim = "" AndAlso _
                       strBed15.Trim = "" AndAlso _
                       strBedE.Trim = "" AndAlso _
                       strBed20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strBedM.Trim

                        '10
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBed10.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBedL.Trim

                        '15
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBed15.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBedE.Trim

                        '20
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBed20.Trim

                    End If

                Else

                    ''全てが空白なら
                    If strBedM.Trim = "" AndAlso _
                       strBedL.Trim = "" AndAlso _
                       strBedE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝
                        strRtn = strBedM.Trim

                        '昼
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBedL.Trim

                        '夕
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strBedE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfSelEtDivDefM
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
    Private Function pfGetSelEtDivDefM(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                       ByVal SysInf As Rec_XSys, _
                                       ByVal blnSnack10 As Boolean, _
                                       ByVal blnSnack15 As Boolean, _
                                       ByVal blnSnack20 As Boolean) As String
        Dim strRtn As String

        Try

            If objRecOrderSTD.SelEtDivDef <> -1 Then

                ''選択食（デフォルト）が入力された場合
                strRtn = CType(objRecOrderSTD.SelEtDivDef, String)

            Else

                Dim strSelEtDivDefM As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivDef10 As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivDefL As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivDef15 As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivDefE As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivDef20 As String = CType(ClsCst.enmSelEtDiv.A, String)

                Dim strSelEtDivMChg As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDiv10Chg As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivLChg As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDiv15Chg As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDivEChg As String = CType(ClsCst.enmSelEtDiv.A, String)
                Dim strSelEtDiv20Chg As String = CType(ClsCst.enmSelEtDiv.A, String)

                '朝
                If objRecOrderSTD.SelEtDivDefM = -1 Then
                    strSelEtDivDefM = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDefM = CType(objRecOrderSTD.SelEtDivDefM, String)
                End If

                '10時
                If objRecOrderSTD.SelEtDivDef10 = -1 Then
                    strSelEtDivDef10 = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDef10 = CType(objRecOrderSTD.SelEtDivDef10, String)
                End If

                '昼
                If objRecOrderSTD.SelEtDivDefL = -1 Then
                    strSelEtDivDefL = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDefL = CType(objRecOrderSTD.SelEtDivDefL, String)
                End If

                '15時
                If objRecOrderSTD.SelEtDivDef15 = -1 Then
                    strSelEtDivDef15 = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDef15 = CType(objRecOrderSTD.SelEtDivDef15, String)
                End If

                '夕
                If objRecOrderSTD.SelEtDivDefE = -1 Then
                    strSelEtDivDefE = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDefE = CType(objRecOrderSTD.SelEtDivDefE, String)
                End If

                '20時
                If objRecOrderSTD.SelEtDivDef20 = -1 Then
                    strSelEtDivDef20 = CType(ClsCst.enmSelEtDiv.A, String)
                Else
                    strSelEtDivDef20 = CType(objRecOrderSTD.SelEtDivDef20, String)
                End If

                ''主食変換がある場合
                If SysInf.OrderBreadSelVal <> "" Then

                    Dim intPipeCnt As Integer

                    intPipeCnt = pfGetIndexCountFromCollection(SysInf.OrderBreadSelVal)

                    Select Case intPipeCnt

                        Case 1

                            strSelEtDivMChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDiv10Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDivLChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDiv15Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDivEChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDiv20Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)

                        Case 3

                            strSelEtDivMChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDivLChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 1, strSelEtDivDefM)
                            strSelEtDivEChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 2, strSelEtDivDefM)

                        Case 6

                            strSelEtDivMChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 0, strSelEtDivDefM)
                            strSelEtDiv10Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 1, strSelEtDivDefM)
                            strSelEtDivLChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 2, strSelEtDivDefM)
                            strSelEtDiv15Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 3, strSelEtDivDefM)
                            strSelEtDivEChg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 4, strSelEtDivDefM)
                            strSelEtDiv20Chg = pfGetIndexStringFromCollection(SysInf.OrderBreadSelVal, 5, strSelEtDivDefM)

                        Case Else

                            ''変換なし
                            strSelEtDivMChg = strSelEtDivDefM
                            strSelEtDiv10Chg = strSelEtDivDef10
                            strSelEtDivLChg = strSelEtDivDefL
                            strSelEtDiv15Chg = strSelEtDivDef15
                            strSelEtDivEChg = strSelEtDivDefE
                            strSelEtDiv20Chg = strSelEtDivDef20

                    End Select

                Else

                    strSelEtDivMChg = strSelEtDivDefM
                    strSelEtDiv10Chg = strSelEtDivDef10
                    strSelEtDivLChg = strSelEtDivDefL
                    strSelEtDiv15Chg = strSelEtDivDef15
                    strSelEtDivEChg = strSelEtDivDefE
                    strSelEtDiv20Chg = strSelEtDivDef20

                End If

                '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                '／主食コードによる選択食変換を行う　　　　　　　　　　　　　＼

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    '朝
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdM) = True Then
                        strRtn = strSelEtDivMChg
                    Else
                        strRtn = strSelEtDivDefM
                    End If

                    ''10時
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCd10) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDiv10Chg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDef10
                    End If

                    ''昼
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdL) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivLChg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDefL
                    End If

                    ''15時
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCd15) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDiv15Chg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDef15
                    End If

                    ''夕
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdE) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivEChg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDefE
                    End If

                    ''20時
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCd20) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDiv20Chg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDef20
                    End If

                Else

                    '朝
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdM) = True Then
                        strRtn = strSelEtDivMChg
                    Else
                        strRtn = strSelEtDivDefM
                    End If

                    ''昼
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdL) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivLChg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDefL
                    End If

                    ''夕
                    If CheckBreadFdCd(SysInf, objRecOrderSTD.MaCkCdE) = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivEChg
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSelEtDivDefE
                    End If

                End If

                ' ''間食がある場合
                'If blnSnack10 = True OrElse _
                '   blnSnack15 = True OrElse _
                '   blnSnack20 = True Then

                '    '朝
                '    If objRecOrderSTD.SelEtDivDefM = -1 Then
                '        strRtn = CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn = CType(objRecOrderSTD.SelEtDivDefM, String)
                '    End If

                '    '10時
                '    If objRecOrderSTD.SelEtDivDef10 = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDef10, String)
                '    End If

                '    '昼
                '    If objRecOrderSTD.SelEtDivDefL = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDefL, String)
                '    End If

                '    '15時
                '    If objRecOrderSTD.SelEtDivDef15 = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDef15, String)
                '    End If

                '    '夕
                '    If objRecOrderSTD.SelEtDivDefE = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDefE, String)
                '    End If

                '    '20時
                '    If objRecOrderSTD.SelEtDivDef20 = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDef20, String)
                '    End If

                'Else

                '    '朝
                '    If objRecOrderSTD.SelEtDivDefM = -1 Then
                '        strRtn = CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn = CType(objRecOrderSTD.SelEtDivDefM, String)
                '    End If

                '    '昼
                '    If objRecOrderSTD.SelEtDivDefL = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDefL, String)
                '    End If

                '    '夕
                '    If objRecOrderSTD.SelEtDivDefE = -1 Then
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmSelEtDiv.A, String)
                '    Else
                '        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.SelEtDivDefE, String)
                '    End If

                'End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetMCPrt
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
    Private Function pfGetMCPrt(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                ByVal objHashMeal As SortedList, _
                                ByVal objMealMHash As Hashtable, _
                                ByVal blnSnack10 As Boolean, _
                                ByVal blnSnack15 As Boolean, _
                                ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String
        Dim strWkMealCd As String = ""

        Try

            If objRecOrderSTD.MCPrt <> -1 Then

                ''欠食が入力された場合
                strRtn = CType(objRecOrderSTD.MCPrt, String)

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    '朝
                    If objRecOrderSTD.MCPrtM = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn = CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn = CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn = CType(objRecOrderSTD.MCPrtM, String)
                    End If

                    '10時
                    If objRecOrderSTD.MCPrt10 = -1 Then
                        '朝
                        If objRecOrderSTD.MCPrtM = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtM, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrt10, String)
                    End If

                    '昼
                    If objRecOrderSTD.MCPrtL = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtL, String)
                    End If

                    '15時
                    If objRecOrderSTD.MCPrt15 = -1 Then
                        '昼
                        If objRecOrderSTD.MCPrtL = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                            End If

                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtL, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrt15, String)
                    End If

                    '夕
                    If objRecOrderSTD.MCPrtE = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtE, String)
                    End If

                    '20時
                    If objRecOrderSTD.MCPrt20 = -1 Then
                        '夕
                        If objRecOrderSTD.MCPrtE = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                            End If

                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtE, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrt20, String)
                    End If

                Else

                    '朝
                    If objRecOrderSTD.MCPrtM = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn = CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn = CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn = CType(objRecOrderSTD.MCPrtM, String)
                    End If

                    '昼
                    If objRecOrderSTD.MCPrtL = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtL, String)
                    End If

                    '夕
                    If objRecOrderSTD.MCPrtE = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MCPrtDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmMCPrt.None, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.MCPrtE, String)
                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetEatAdDiv
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
    Private Function pfGetEatAdDiv(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                   ByVal objHashMeal As SortedList, _
                                   ByVal objMealMHash As Hashtable, _
                                   ByVal blnSnack10 As Boolean, _
                                   ByVal blnSnack15 As Boolean, _
                                   ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String
        Dim strWkMealCd As String = ""

        Try

            '1日加算の場合
            If objRecOrderSTD.EatAdDiv <> -1 Then

                ''加算が入力された場合
                strRtn = CType(objRecOrderSTD.EatAdDiv, String)

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    '朝
                    If objRecOrderSTD.EatAdDivM = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn = CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn = CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If

                    Else
                        strRtn = CType(objRecOrderSTD.EatAdDivM, String)
                    End If

                    '10時
                    If objRecOrderSTD.EatAdDiv10 = -1 Then
                        '朝
                        If objRecOrderSTD.EatAdDivM = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                            End If

                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivM, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDiv10, String)
                    End If

                    '昼
                    If objRecOrderSTD.EatAdDivL = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivL, String)
                    End If

                    '15時
                    If objRecOrderSTD.EatAdDiv15 = -1 Then
                        '昼
                        If objRecOrderSTD.EatAdDivL = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                            End If

                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivL, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDiv15, String)
                    End If

                    '夕
                    If objRecOrderSTD.EatAdDivE = -1 Then

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If

                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivE, String)
                    End If

                    '20時
                    If objRecOrderSTD.EatAdDiv20 = -1 Then
                        '夕
                        If objRecOrderSTD.EatAdDivE = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                            End If

                            If objMealMHash.ContainsKey(strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                            End If

                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivE, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDiv20, String)
                    End If

                Else

                    '朝
                    If objRecOrderSTD.EatAdDivM = -1 Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn = CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn = CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If
                    Else
                        strRtn = CType(objRecOrderSTD.EatAdDivM, String)
                    End If

                    '昼
                    If objRecOrderSTD.EatAdDivL = -1 Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivL, String)
                    End If

                    '夕
                    If objRecOrderSTD.EatAdDivE = -1 Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).EatAdDiv, String)
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmEatAdDiv.NoAdd, String)
                        End If
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.EatAdDivE, String)
                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNoMlRs
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
    Private Function pfGetSdShap(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objHashMeal As SortedList, _
                                 ByVal objMealMHash As Hashtable, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String
        Dim strWkMealCd As String = ""

        Try

            strRtn = ""

            Dim strSdShap As String
            Dim strSdShapM As String
            Dim strSdShap10 As String
            Dim strSdShapL As String
            Dim strSdShap15 As String
            Dim strSdShapE As String
            Dim strSdShap20 As String

            ''コードを取得と変換
            strSdShap = pfChangSdShap(objRecOrderSTD.SdShap.Trim)
            strSdShapM = pfChangSdShap(objRecOrderSTD.SdShapM.Trim)
            strSdShap10 = pfChangSdShap(objRecOrderSTD.SdShap10.Trim)
            strSdShapL = pfChangSdShap(objRecOrderSTD.SdShapL.Trim)
            strSdShap15 = pfChangSdShap(objRecOrderSTD.SdShap15.Trim)
            strSdShapE = pfChangSdShap(objRecOrderSTD.SdShapE.Trim)
            strSdShap20 = pfChangSdShap(objRecOrderSTD.SdShap20.Trim)

            '1日の形態の場合
            If strSdShap.Trim <> "" Then

                strRtn = strSdShap.Trim

            Else

                ''間がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''空白なら食種マスタを見る（朝）
                    If strSdShapM.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapM = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（10時）
                    If strSdShap10.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.AM10Div) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.AM10Div), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShap10 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（昼）
                    If strSdShapL.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapL = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（15時）
                    If strSdShap15.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.PM15Div) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.PM15Div), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShap15 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（夕）
                    If strSdShapE.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapE = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（20時）
                    If strSdShap20.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.NightDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.NightDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShap20 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''全て空白なら
                    If strSdShapM.Trim = "" AndAlso strSdShap10.Trim = "" AndAlso _
                       strSdShapL.Trim = "" AndAlso strSdShap15.Trim = "" AndAlso _
                       strSdShapE.Trim = "" AndAlso strSdShap20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝の形態
                        strRtn = strSdShapM.Trim

                        '10の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShap10.Trim

                        '昼の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShapL.Trim

                        '15の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShap15.Trim

                        '夕の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShapE.Trim

                        '20の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShap20.Trim

                    End If

                Else

                    ''空白なら食種マスタを見る（朝）
                    If strSdShapM.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapM = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（昼）
                    If strSdShapL.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapL = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''空白なら食種マスタを見る（夕）
                    If strSdShapE.Trim = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strSdShapE = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).SdShape
                        End If
                    End If

                    ''全て空白なら
                    If strSdShapM.Trim = "" AndAlso strSdShapL.Trim = "" AndAlso _
                       strSdShapE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝の形態
                        strRtn = strSdShapM.Trim

                        '昼の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShapL.Trim

                        '夕の形態
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strSdShapE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfGetMaCkCd
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
    Private Function pfGetMaCkCd(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objRecOMain As Rec_OMain, _
                                 ByVal objHashMeal As SortedList, _
                                 ByVal objMealMHash As Hashtable, _
                                 ByVal strAddVal As String, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String = ""
        Dim strWkMealCd As String = ""

        Try

            strRtn = ""

            Dim strMaCkCd As String = ""
            Dim strMaCkCdM As String = ""
            Dim strMaCkCd10 As String = ""
            Dim strMaCkCdL As String = ""
            Dim strMaCkCd15 As String = ""
            Dim strMaCkCdE As String = ""
            Dim strMaCkCd20 As String = ""


            ''コードを取得と変換
            strMaCkCd = pfChangMaCkCd(objRecOrderSTD.MaCkCd.Trim)
            strMaCkCdM = pfChangMaCkCd(objRecOrderSTD.MaCkCdM.Trim)
            strMaCkCd10 = pfChangMaCkCd(objRecOrderSTD.MaCkCd10.Trim)
            strMaCkCdL = pfChangMaCkCd(objRecOrderSTD.MaCkCdL.Trim)
            strMaCkCd15 = pfChangMaCkCd(objRecOrderSTD.MaCkCd15.Trim)
            strMaCkCdE = pfChangMaCkCd(objRecOrderSTD.MaCkCdE.Trim)
            strMaCkCd20 = pfChangMaCkCd(objRecOrderSTD.MaCkCd20.Trim)

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／主食変換処理を行う　　　　　　　　　　　　　　　　　　　　＼
            If objRecOMain.MaCkCdMNew.Trim <> "" Then
                strMaCkCdM = objRecOMain.MaCkCdMNew
                strMaCkCd10 = objRecOMain.MaCkCdMNew
                strMaCkCd = ""
            End If
            If objRecOMain.MaCkCdLNew.Trim <> "" Then
                strMaCkCdL = objRecOMain.MaCkCdLNew
                strMaCkCd15 = objRecOMain.MaCkCdLNew
                strMaCkCd = ""
            End If
            If objRecOMain.MaCkCdENew.Trim <> "" Then
                strMaCkCdE = objRecOMain.MaCkCdENew
                strMaCkCd20 = objRecOMain.MaCkCdENew
                strMaCkCd = ""
            End If

            If objRecOMain.MaCkCdMNew.Trim <> "" OrElse _
               objRecOMain.MaCkCdLNew.Trim <> "" OrElse _
               objRecOMain.MaCkCdENew.Trim <> "" Then

                ''変換後は全て同じなら
                If objRecOMain.MaCkCdMNew = objRecOMain.MaCkCdLNew AndAlso _
                   objRecOMain.MaCkCdMNew = objRecOMain.MaCkCdENew Then
                    strMaCkCd = objRecOMain.MaCkCdMNew
                End If

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／食種が変換されている場合　　　　　　　　　　　　　　　　　＼
            If objRecOMain.MealCdMNew.Trim = "" AndAlso _
               objRecOMain.MealCdLNew.Trim = "" AndAlso _
               objRecOMain.MealCdENew.Trim = "" Then

                ''食種変換なし なにもしない

            Else

                ''食種変換あり でもなにもしない


            End If

            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '1日の主食の場合
            If strMaCkCd <> "" Then

                strRtn = strAddVal & strMaCkCd

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    ''空白なら食種マスタの主食を見る（朝）
                    If strMaCkCdM = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdM = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkMorn
                        End If
                    Else
                        strMaCkCdM = strAddVal & strMaCkCdM
                    End If

                    ''空白なら食種マスタの主食を見る（10時）
                    If strMaCkCd10 = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.AM10Div) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.AM10Div), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCd10 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkMorn
                        End If
                    Else
                        strMaCkCd10 = strAddVal & strMaCkCd10
                    End If

                    ''空白なら食種マスタの主食を見る（昼）
                    If strMaCkCdL = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdL = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkLun
                        End If
                    Else
                        strMaCkCdL = strAddVal & strMaCkCdL
                    End If

                    ''空白なら食種マスタの主食を見る（15時）
                    If strMaCkCd15 = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.PM15Div) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.PM15Div), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCd15 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkLun
                        End If
                    Else
                        strMaCkCd15 = strAddVal & strMaCkCd15
                    End If

                    ''空白なら食種マスタの主食を見る（夕）
                    If strMaCkCdE = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdE = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkEvo
                        End If
                    Else
                        strMaCkCdE = strAddVal & strMaCkCdE
                    End If

                    ''空白なら食種マスタの主食を見る（20時）
                    If strMaCkCd20 = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.NightDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.NightDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCd20 = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkEvo
                        End If
                    Else
                        strMaCkCd20 = strAddVal & strMaCkCd20
                    End If

                    ''全て空白なら
                    If strMaCkCdM.Trim = "" AndAlso strMaCkCd10.Trim = "" AndAlso _
                       strMaCkCdL.Trim = "" AndAlso strMaCkCd15.Trim = "" AndAlso _
                       strMaCkCdE.Trim = "" AndAlso strMaCkCd20.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝の主食
                        strRtn = strMaCkCdM.Trim
                        '10の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCd10.Trim
                        '昼の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCdL.Trim
                        '15の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCd15.Trim
                        '夕の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCdE.Trim
                        '20の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCd20.Trim

                    End If

                Else

                    ''空白なら食種マスタの主食を見る（朝）
                    If strMaCkCdM = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdM = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkMorn
                        End If
                    Else
                        strMaCkCdM = strAddVal & strMaCkCdM
                    End If

                    ''空白なら食種マスタの主食を見る（昼）
                    If strMaCkCdL = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdL = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkLun
                        End If
                    Else
                        strMaCkCdL = strAddVal & strMaCkCdL
                    End If

                    ''空白なら食種マスタの主食を見る（夕）
                    If strMaCkCdE = "" Then
                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                        End If

                        If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                            strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                        End If

                        If objMealMHash.ContainsKey(strWkMealCd) = True Then
                            strMaCkCdE = CType(objMealMHash.Item(strWkMealCd), Rec_MMealTpP).MaCkEvo
                        End If
                    Else
                        strMaCkCdE = strAddVal & strMaCkCdE
                    End If

                    ''全て空白なら
                    If strMaCkCdM.Trim = "" AndAlso strMaCkCdL.Trim = "" AndAlso _
                       strMaCkCdE.Trim = "" Then

                        strRtn = ""

                    Else

                        '朝の主食
                        strRtn = strMaCkCdM.Trim
                        '昼の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCdL.Trim
                        '夕の主食
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & strMaCkCdE.Trim

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function
    '**********************************************************************
    ' Name       :pfGetNoMlRs
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
    Private Function pfGetMealCd(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objRecOMain As Rec_OMain, _
                                 ByVal SysInf As Rec_XSys, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As SortedList

        Dim objMealCd As SortedList
        Dim blnChangeNoMeal As Boolean

        Try

            Dim strMealCd As String = ""
            Dim strMealCdM As String = ""
            Dim strMealCd10 As String = ""
            Dim strMealCdL As String = ""
            Dim strMealCd15 As String = ""
            Dim strMealCdE As String = ""
            Dim strMealCd20 As String = ""

            ''コードを取得と変換
            strMealCd = pfChangMealCd(objRecOrderSTD.MealCd.Trim)
            strMealCdM = pfChangMealCd(objRecOrderSTD.MealCdM.Trim)
            strMealCd10 = pfChangMealCd(objRecOrderSTD.MealCd10.Trim)
            strMealCdL = pfChangMealCd(objRecOrderSTD.MealCdL.Trim)
            strMealCd15 = pfChangMealCd(objRecOrderSTD.MealCd15.Trim)
            strMealCdE = pfChangMealCd(objRecOrderSTD.MealCdE.Trim)
            strMealCd20 = pfChangMealCd(objRecOrderSTD.MealCd20.Trim)

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／主食コードによる欠食変換を行う　　　　　　　　　　　　　　＼
            blnChangeNoMeal = False

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCd) = True Then
                strMealCd = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCd)
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCdM) = True Then
                strMealCdM = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCdM)
                blnChangeNoMeal = True
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCd10) = True Then
                strMealCd10 = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCd10)
                blnChangeNoMeal = True
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCdL) = True Then
                strMealCdL = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCdL)
                blnChangeNoMeal = True
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCd15) = True Then
                strMealCd15 = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCd15)
                blnChangeNoMeal = True
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCdE) = True Then
                strMealCdE = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCdE)
                blnChangeNoMeal = True
            End If

            If CheckNoMainFdCd(SysInf, objRecOrderSTD.MaCkCd20) = True Then
                strMealCd20 = pfGetFirstStringFromCollection(SysInf.OrderNoMealCd, strMealCd20)
                blnChangeNoMeal = True
            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／


            ''朝・昼・夕で1回でも変換があった場合、
            If blnChangeNoMeal = True Then

                If strMealCdM = "" Then
                    strMealCdM = strMealCd
                End If

                If strMealCd10 = "" Then
                    If strMealCdM = "" Then
                        strMealCd10 = strMealCd
                    Else
                        strMealCd10 = strMealCdM
                    End If
                End If

                If strMealCdL = "" Then
                    strMealCdL = strMealCd
                End If

                If strMealCd15 = "" Then
                    If strMealCdL = "" Then
                        strMealCd15 = strMealCd
                    Else
                        strMealCd15 = strMealCdL
                    End If
                End If

                If strMealCdE = "" Then
                    strMealCdE = strMealCd
                End If

                If strMealCd20 = "" Then
                    If strMealCdE = "" Then
                        strMealCd20 = strMealCd
                    Else
                        strMealCd20 = strMealCdE
                    End If
                End If

                strMealCd = ""

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／食種変換処理を行う　　　　　　　　　　　　　　　　　　　　＼
            blnChangeNoMeal = False

            If objRecOMain.MealCdMNew.Trim <> "" Then
                strMealCdM = objRecOMain.MealCdMNew
                strMealCd10 = objRecOMain.MealCdMNew
                blnChangeNoMeal = True
            End If
            If objRecOMain.MealCdLNew.Trim <> "" Then
                strMealCdL = objRecOMain.MealCdLNew
                strMealCd15 = objRecOMain.MealCdLNew
                blnChangeNoMeal = True
            End If
            If objRecOMain.MealCdENew.Trim <> "" Then
                strMealCdE = objRecOMain.MealCdENew
                strMealCd20 = objRecOMain.MealCdENew
                blnChangeNoMeal = True
            End If

            If blnChangeNoMeal = True Then

                If strMealCdM = "" Then
                    strMealCdM = strMealCd
                End If

                If strMealCd10 = "" Then
                    If strMealCdM = "" Then
                        strMealCd10 = strMealCd
                    Else
                        strMealCd10 = strMealCdM
                    End If
                End If

                If strMealCdL = "" Then
                    strMealCdL = strMealCd
                End If

                If strMealCd15 = "" Then
                    If strMealCdL = "" Then
                        strMealCd15 = strMealCd
                    Else
                        strMealCd15 = strMealCdL
                    End If
                End If

                If strMealCdE = "" Then
                    strMealCdE = strMealCd
                End If

                If strMealCd20 = "" Then
                    If strMealCdE = "" Then
                        strMealCd20 = strMealCd
                    Else
                        strMealCd20 = strMealCdE
                    End If
                End If

                If objRecOMain.MealCdMNew.Trim <> "" OrElse _
                   objRecOMain.MealCdLNew.Trim <> "" OrElse _
                   objRecOMain.MealCdENew.Trim <> "" Then

                    ''変換後は全て同じなら
                    If objRecOMain.MealCdMNew = objRecOMain.MealCdLNew AndAlso _
                       objRecOMain.MealCdMNew = objRecOMain.MealCdENew Then
                        strMealCd = objRecOMain.MealCdMNew
                    End If

                End If

                strMealCd = ""

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／病棟による食種変換処理を行う　　　　　　　　　　　　　　　＼
            If SysInf.OrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderWardCnvSet) <> -1 Then

                Dim strWardCgCd As String = ""
                Dim strWardCgCdM As String = ""
                Dim strWardCgCd10 As String = ""
                Dim strWardCgCdL As String = ""
                Dim strWardCgCd15 As String = ""
                Dim strWardCgCdE As String = ""
                Dim strWardCgCd20 As String = ""

                ''病棟ごとの変換コードを取得で検索
                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.Ward.Trim) = True Then
                    strWardCgCd = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.Ward.Trim), String)
                    strWardCgCdM = strWardCgCd
                    strWardCgCd10 = strWardCgCd
                    strWardCgCdL = strWardCgCd
                    strWardCgCd15 = strWardCgCd
                    strWardCgCdE = strWardCgCd
                    strWardCgCd20 = strWardCgCd
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.WardM.Trim) = True Then
                    strWardCgCdM = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.WardM.Trim), String)
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.Ward10.Trim) = True Then
                    strWardCgCd10 = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.Ward10.Trim), String)
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.WardL.Trim) = True Then
                    strWardCgCdL = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.WardL.Trim), String)
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.Ward15.Trim) = True Then
                    strWardCgCd15 = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.Ward15.Trim), String)
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.WardE.Trim) = True Then
                    strWardCgCdE = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.WardE.Trim), String)
                End If

                If m_hashWardMealAddCd.ContainsKey(objRecOrderSTD.Ward20.Trim) = True Then
                    strWardCgCd20 = CType(m_hashWardMealAddCd.Item(objRecOrderSTD.Ward20.Trim), String)
                End If

                If strMealCd <> "" Then

                    '1日の食種の場合
                    If IsNumeric(strMealCd) = True AndAlso IsNumeric(strWardCgCd) = True Then
                        strMealCd = CType(CType(strMealCd, Integer) + CType(strWardCgCd, Integer), String)
                        strMealCd = strMealCd.PadLeft(4, "0"c)
                    End If

                Else

                    ''複数食種の場合
                    If IsNumeric(strMealCdM) = True AndAlso IsNumeric(strWardCgCdM) = True Then
                        strMealCdM = CType(CType(strMealCdM, Integer) + CType(strWardCgCdM, Integer), String)
                        strMealCdM = strMealCdM.PadLeft(4, "0"c)
                    End If

                    If IsNumeric(strMealCd10) = True AndAlso IsNumeric(strWardCgCd10) = True Then
                        strMealCd10 = CType(CType(strMealCd10, Integer) + CType(strWardCgCd10, Integer), String)
                        strMealCd10 = strMealCd10.PadLeft(4, "0"c)
                    End If

                    If IsNumeric(strMealCdL) = True AndAlso IsNumeric(strWardCgCdL) = True Then
                        strMealCdL = CType(CType(strMealCdL, Integer) + CType(strWardCgCdL, Integer), String)
                        strMealCdL = strMealCdL.PadLeft(4, "0"c)
                    End If

                    If IsNumeric(strMealCd15) = True AndAlso IsNumeric(strWardCgCd15) = True Then
                        strMealCd15 = CType(CType(strMealCd15, Integer) + CType(strWardCgCd15, Integer), String)
                        strMealCd15 = strMealCd15.PadLeft(4, "0"c)
                    End If

                    If IsNumeric(strMealCdE) = True AndAlso IsNumeric(strWardCgCdE) = True Then
                        strMealCdE = CType(CType(strMealCdE, Integer) + CType(strWardCgCdE, Integer), String)
                        strMealCdE = strMealCdE.PadLeft(4, "0"c)
                    End If

                    If IsNumeric(strMealCd20) = True AndAlso IsNumeric(strWardCgCd20) = True Then
                        strMealCd20 = CType(CType(strMealCd20, Integer) + CType(strWardCgCd20, Integer), String)
                        strMealCd20 = strMealCd20.PadLeft(4, "0"c)
                    End If

                End If

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            objMealCd = New SortedList

            If strMealCd <> "" Then

                objMealCd.Add(ClsCst.enmEatTimeDiv.OneDay, strMealCd)

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    '朝の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.MorningDiv, strMealCdM)

                    '10の食種
                    If strMealCd10 = "" Then
                        objMealCd.Add(ClsCst.enmEatTimeDiv.AM10Div, strMealCdM)
                    Else
                        objMealCd.Add(ClsCst.enmEatTimeDiv.AM10Div, strMealCd10)
                    End If

                    '昼の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.LunchDiv, strMealCdL)

                    '15の食種
                    If strMealCd15 = "" Then
                        objMealCd.Add(ClsCst.enmEatTimeDiv.PM15Div, strMealCdL)
                    Else
                        objMealCd.Add(ClsCst.enmEatTimeDiv.PM15Div, strMealCd15)
                    End If

                    '夕の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.EveningDiv, strMealCdE)

                    '20の食種
                    If strMealCd20 = "" Then
                        objMealCd.Add(ClsCst.enmEatTimeDiv.NightDiv, strMealCdE)
                    Else
                        objMealCd.Add(ClsCst.enmEatTimeDiv.NightDiv, strMealCd20)
                    End If

                Else

                    '朝の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.MorningDiv, strMealCdM)
                    '昼の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.LunchDiv, strMealCdL)
                    '夕の食種
                    objMealCd.Add(ClsCst.enmEatTimeDiv.EveningDiv, strMealCdE)

                End If

            End If

            Return objMealCd

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNoMlRs
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
    Private Function pfGetNoMlRs(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objHashMeal As SortedList, _
                                 ByVal SysInf As Rec_XSys, _
                                 ByVal strComOdDiv As String, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String = ""

        Dim strNoMlRsM As String = ""
        Dim strNoMlRs10 As String = ""
        Dim strNoMlRsL As String = ""
        Dim strNoMlRs15 As String = ""
        Dim strNoMlRsE As String = ""
        Dim strNoMlRs20 As String = ""
        Dim strWkMealCd As String = ""

        Try

            ''欠食オーダ時は無条件で欠食とする
            If strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMeal OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMealInHospital OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMealMove Then

                strRtn = SysInf.OrderNoMlRs     'デフォルトの欠食理由

            ElseIf strComOdDiv = ClsCst.ORecv.OrderComOdDiv_SleepOut OrElse _
                   strComOdDiv = ClsCst.ORecv.OrderComOdDiv_GoOut Then

                If SysInf.OrderGoOutNoMlRs <> "" Then
                    strRtn = SysInf.OrderGoOutNoMlRs    '外泊の欠食理由
                Else
                    strRtn = SysInf.OrderNoMlRs         'デフォルトの欠食理由
                End If

            Else

                '1日欠食の場合
                If objRecOrderSTD.NoMeal = ClsCst.PerCST.enmNoMeal.NoMeal Then

                    If objRecOrderSTD.NoMlRs.Trim = "" Then

                        strRtn = SysInf.OrderNoMlRs  'デフォルトの欠食理由

                    Else
                        ''欠食理由が入力された場合
                        strRtn = objRecOrderSTD.NoMlRs
                    End If

                ElseIf objRecOrderSTD.NoMeal = ClsCst.PerCST.enmNoMeal.Normal Then

                    ''欠食ではない
                    strRtn = ""

                Else

                    ''1日欠食ではない場合

                    ''間食がある場合
                    If blnSnack10 = True OrElse _
                       blnSnack15 = True OrElse _
                       blnSnack20 = True Then

                        strRtn = ""

                        '朝
                        Select Case objRecOrderSTD.NoMealM
                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                                End If

                                strNoMlRsM = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                                '欠食コメント
                                If SysInf.NoMealCmtMon.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdM
                                            If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If


                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsM = ""

                                '欠食コメント
                                If SysInf.NoMealCmtMon.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdM
                                            If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsM.Trim = "" Then
                                    strNoMlRsM = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsM = objRecOrderSTD.NoMlRsM
                                End If

                        End Select

                        '10時
                        Select Case objRecOrderSTD.NoMeal10

                            Case -1

                                ''入力なし
                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.AM10Div) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.AM10Div), String)
                                End If

                                strNoMlRs10 = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由


                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRs10 = ""

                                '欠食コメント
                                If SysInf.NoMealCmtAM10.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtAM10, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRs10 = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd10
                                            If SysInf.NoMealCmtAM10.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtAM10, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRs10 = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRs10.Trim = "" Then
                                    strNoMlRs10 = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRs10 = objRecOrderSTD.NoMlRs10
                                End If

                        End Select

                        '昼
                        Select Case objRecOrderSTD.NoMealL

                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                                End If

                                strNoMlRsL = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                                '欠食コメント
                                If SysInf.NoMealCmtLun.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdL
                                            If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsL = ""

                                '欠食コメント
                                If SysInf.NoMealCmtLun.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdL
                                            If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsL.Trim = "" Then
                                    strNoMlRsL = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsL = objRecOrderSTD.NoMlRsL
                                End If

                        End Select

                        '15時
                        Select Case objRecOrderSTD.NoMeal15

                            Case -1

                                ''入力なし
                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.PM15Div) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.PM15Div), String)
                                End If

                                strNoMlRs15 = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRs15 = ""

                                '欠食コメント
                                If SysInf.NoMealCmtPM15.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtPM15, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRs15 = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd15
                                            If SysInf.NoMealCmtPM15.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtPM15, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRs15 = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRs15.Trim = "" Then
                                    strNoMlRs15 = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRs15 = objRecOrderSTD.NoMlRs15
                                End If

                        End Select

                        '夕
                        Select Case objRecOrderSTD.NoMealE

                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                                End If

                                strNoMlRsE = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                                '欠食コメント
                                If SysInf.NoMealCmtEve.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdE
                                            If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsE = ""

                                '欠食コメント
                                If SysInf.NoMealCmtEve.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdE
                                            If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsE.Trim = "" Then
                                    strNoMlRsE = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsE = objRecOrderSTD.NoMlRsE
                                End If

                        End Select

                        '20時
                        Select Case objRecOrderSTD.NoMeal20

                            Case -1

                                ''入力なし
                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.NightDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.NightDiv), String)
                                End If

                                strNoMlRs20 = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRs20 = ""


                                '欠食コメント
                                If SysInf.NoMealCmtPM20.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtPM20, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                    strNoMlRs20 = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd20
                                            If SysInf.NoMealCmtPM20.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtPM20, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 AndAlso strNoMealCmt.Split(":"c).Length = 2 Then
                                                        strNoMlRs20 = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRs20.Trim = "" Then
                                    strNoMlRs20 = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRs20 = objRecOrderSTD.NoMlRs20
                                End If

                        End Select

                        ''全て空白なら
                        If strNoMlRsM.Trim = "" AndAlso strNoMlRs10.Trim = "" AndAlso _
                           strNoMlRsL.Trim = "" AndAlso strNoMlRs15.Trim = "" AndAlso _
                           strNoMlRsE.Trim = "" AndAlso strNoMlRs20.Trim = "" Then

                            strRtn = ""

                        Else

                            strRtn = strNoMlRsM.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRs10.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRsL.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRs15.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRsE.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRs20.Trim

                        End If


                    Else

                        strRtn = ""

                        '朝
                        Select Case objRecOrderSTD.NoMealM
                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                                End If

                                strNoMlRsM = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由


                                '欠食コメント
                                If SysInf.NoMealCmtMon.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                    strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdM
                                            If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtMon, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                        strNoMlRsM = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsM = ""

                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsM.Trim = "" Then
                                    strNoMlRsM = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsM = objRecOrderSTD.NoMlRsM
                                End If
                        End Select

                        '昼
                        Select Case objRecOrderSTD.NoMealL

                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                                End If

                                strNoMlRsL = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由


                                '欠食コメント
                                If SysInf.NoMealCmtLun.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                    strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdL
                                            If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtLun, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                        strNoMlRsL = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsL = ""

                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsL.Trim = "" Then
                                    strNoMlRsL = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsL = objRecOrderSTD.NoMlRsL
                                End If

                        End Select


                        '夕
                        Select Case objRecOrderSTD.NoMealE

                            Case -1

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                                End If

                                strNoMlRsE = GetNoMealRs(SysInf, strWkMealCd)  '欠食理由

                                '欠食コメント
                                If SysInf.NoMealCmtEve.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                            For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                    strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                End If
                                            Next
                                        End If
                                    Next

                                    If strNoMlRsM.Trim = "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCdE
                                            If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                                For Each strNoMealCmt As String In SplitOrderData(SysInf.NoMealCmtEve, ORDER_DELIMITER_CHAR_PIPE)
                                                    If strNoMealCmt.IndexOf(strCmt) >= 0 Then
                                                        strNoMlRsE = strNoMealCmt.Split(":"c)(1)
                                                    End If
                                                Next
                                            End If
                                        Next
                                    End If
                                End If
                            Case ClsCst.PerCST.enmNoMeal.Normal

                                ''欠食ではない
                                strNoMlRsE = ""

                            Case ClsCst.PerCST.enmNoMeal.NoMeal

                                If objRecOrderSTD.NoMlRsE.Trim = "" Then
                                    strNoMlRsE = SysInf.OrderNoMlRs  'デフォルトの欠食理由
                                Else
                                    ''欠食理由が入力された場合
                                    strNoMlRsE = objRecOrderSTD.NoMlRsE
                                End If

                        End Select

                        ''全て空白なら
                        If strNoMlRsM.Trim = "" AndAlso strNoMlRsL.Trim = "" AndAlso _
                           strNoMlRsE.Trim = "" Then

                            strRtn = ""

                        Else

                            strRtn = strNoMlRsM.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRsL.Trim
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & strNoMlRsE.Trim

                        End If

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetValidF
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
    Private Function pfGetValidF(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String

        Try

            If objRecOrderSTD.ValidEatTmDivM = -1 AndAlso _
               objRecOrderSTD.ValidEatTmDivL = -1 AndAlso _
               objRecOrderSTD.ValidEatTmDivE = -1 Then

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    strRtn = "1"        ''朝

                    If blnSnack10 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "0"
                    End If

                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"       '昼

                    If blnSnack15 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "0"
                    End If

                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"       '夕

                    If blnSnack20 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & "0"
                    End If

                Else

                    ''データがない場合は、朝・昼・夕が有効とみなす
                    strRtn = "1|0|1|0|1|0"

                End If

            Else

                ''間食がある場合
                If blnSnack10 = True OrElse _
                   blnSnack15 = True OrElse _
                   blnSnack20 = True Then

                    '朝
                    If objRecOrderSTD.ValidEatTmDivM = -1 Then
                        strRtn = CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn = CType(objRecOrderSTD.ValidEatTmDivM, String)
                    End If

                    '10時
                    If blnSnack10 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.Exists, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    End If

                    '昼
                    If objRecOrderSTD.ValidEatTmDivL = -1 Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.ValidEatTmDivL, String)
                    End If

                    '15時
                    If blnSnack15 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.Exists, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    End If

                    '夕
                    If objRecOrderSTD.ValidEatTmDivE = -1 Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.ValidEatTmDivE, String)
                    End If

                    '20時
                    If blnSnack20 = True Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.Exists, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    End If

                Else

                    '朝
                    If objRecOrderSTD.ValidEatTmDivM = -1 Then
                        strRtn = CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn = CType(objRecOrderSTD.ValidEatTmDivM, String)
                    End If

                    '10時
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)

                    '昼
                    If objRecOrderSTD.ValidEatTmDivL = -1 Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.ValidEatTmDivL, String)
                    End If

                    '15時
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)

                    '夕
                    If objRecOrderSTD.ValidEatTmDivE = -1 Then
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)
                    Else
                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.ValidEatTmDivE, String)
                    End If

                    '20時
                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.enmValidF.None, String)

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetEatAdDiv
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
    Private Function pfGetKCnt(ByVal objRecOrderSTD As Rec_OrderSTD, _
                               ByVal blnSnack10 As Boolean, _
                               ByVal blnSnack15 As Boolean, _
                               ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String = ""

        Try

            ''曜日指定がある場合
            If objRecOrderSTD.KCntWeek <> "" Then

                ''曜日指定の場合
                Select Case objRecOrderSTD.KCntWeek.Length

                    Case 7

                        ''週間入力の場合
                        For intCnt As Integer = 0 To objRecOrderSTD.KCntWeek.Length - 1

                            If strRtn = "" Then
                                strRtn = objRecOrderSTD.KCntWeek.Substring(intCnt, 1)
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & objRecOrderSTD.KCntWeek.Substring(intCnt, 1)
                            End If

                        Next

                    Case 21

                        ''週間＋朝昼夕で入力の場合
                        strRtn = "1"

                    Case 42

                        ''週間入力＋朝昼夕＋間食の場合
                        strRtn = "1"

                    Case Else

                        strRtn = "1"

                End Select


            Else

                '1日の場合
                If objRecOrderSTD.KCnt <> -1 Then

                    ''数字が入力された場合
                    strRtn = CType(objRecOrderSTD.KCnt, String)

                Else

                    ''間食がある場合
                    If blnSnack10 = True OrElse _
                       blnSnack15 = True OrElse _
                       blnSnack20 = True Then

                        '朝
                        If objRecOrderSTD.KCntM = -1 Then
                            ''数字が入力されない
                            strRtn = "1"
                        Else
                            strRtn = CType(objRecOrderSTD.KCntM, String)
                        End If

                        '10時
                        If objRecOrderSTD.KCnt10 = -1 Then
                            '朝
                            If objRecOrderSTD.EatAdDivM = -1 Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntM, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCnt10, String)
                        End If

                        '昼
                        If objRecOrderSTD.KCntL = -1 Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntL, String)
                        End If

                        '15時
                        If objRecOrderSTD.KCnt15 = -1 Then
                            '昼
                            If objRecOrderSTD.KCntL = -1 Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntL, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCnt15, String)
                        End If

                        '夕
                        If objRecOrderSTD.KCntE = -1 Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntE, String)
                        End If

                        '20時
                        If objRecOrderSTD.KCnt20 = -1 Then
                            '夕
                            If objRecOrderSTD.KCntE = -1 Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntE, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCnt20, String)
                        End If

                    Else

                        '朝
                        If objRecOrderSTD.KCntM = -1 Then
                            strRtn = "1"
                        Else
                            strRtn = CType(objRecOrderSTD.KCntM, String)
                        End If

                        '昼
                        If objRecOrderSTD.KCntL = -1 Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntL, String)
                        End If

                        '夕
                        If objRecOrderSTD.KCntE = -1 Then
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & "1"
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.KCntE, String)
                        End If

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNoMeal
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
    Private Function pfGetNoMeal(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                 ByVal objHashMeal As SortedList, _
                                 ByVal SysInf As Rec_XSys, _
                                 ByVal strComOdDiv As String, _
                                 ByVal blnSnack10 As Boolean, _
                                 ByVal blnSnack15 As Boolean, _
                                 ByVal blnSnack20 As Boolean) As String

        Dim strRtn As String = ""
        Dim strWkMealCd As String = ""

        Try

            ''欠食オーダ時は無条件で欠食とする
            If strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMeal OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMealInHospital OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_NoMealMove OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_SleepOut OrElse _
               strComOdDiv = ClsCst.ORecv.OrderComOdDiv_GoOut Then

                strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)

            Else

                If objRecOrderSTD.NoMeal <> -1 Then

                    ''欠食が入力された場合
                    strRtn = CType(objRecOrderSTD.NoMeal, String)

                Else

                    ''間食がある場合
                    If blnSnack10 = True OrElse _
                       blnSnack15 = True OrElse _
                       blnSnack20 = True Then

                        '朝
                        If objRecOrderSTD.NoMealM = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                strRtn = CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                            End If

                            '欠食コメントがあるかチェック
                            If SysInf.NoMealCmtMon.Trim <> "" Then
                                For Each strCmt As String In objRecOrderSTD.CmtCd
                                    If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                        strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                    End If
                                Next
                            End If

                            If SysInf.NoMealCmtMon.Trim <> "" Then
                                For Each strCmt As String In objRecOrderSTD.CmtCdM
                                    If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                        strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                    End If
                                Next
                            End If
                        Else
                            strRtn = CType(objRecOrderSTD.NoMealM, String)
                        End If

                        '10時
                        If objRecOrderSTD.NoMeal10 = -1 Then
                            '朝
                            If objRecOrderSTD.NoMealM = -1 Then

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.AM10Div) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.AM10Div), String)
                                End If

                                ''欠食かどうかチェック
                                If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                Else
                                    Dim blnFlg As Boolean = False

                                    '欠食コメントがあるかチェック
                                    If SysInf.NoMealCmtAM10.Trim <> "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd
                                            If SysInf.NoMealCmtAM10.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If SysInf.NoMealCmtAM10.Trim <> "" AndAlso blnFlg = False Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd10
                                            If SysInf.NoMealCmtAM10.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If blnFlg = False Then
                                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                    End If
                                End If


                            Else
                                strRtn = CType(objRecOrderSTD.NoMealM, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMeal10, String)
                        End If

                        '昼
                        If objRecOrderSTD.NoMealL = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                Dim blnFlg As Boolean = False

                                '欠食コメントがあるかチェック
                                If SysInf.NoMealCmtLun.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If SysInf.NoMealCmtLun.Trim <> "" AndAlso blnFlg = False Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCdL
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If blnFlg = False Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                End If
                            End If



                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealL, String)
                        End If

                        '15時
                        If objRecOrderSTD.NoMeal15 = -1 Then
                            '昼
                            If objRecOrderSTD.NoMealL = -1 Then

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.PM15Div) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.PM15Div), String)
                                End If

                                ''欠食かどうかチェック
                                If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                Else
                                    Dim blnFlg As Boolean = False
                                    '欠食コメントがあるかチェック
                                    If SysInf.NoMealCmtPM15.Trim <> "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd
                                            If SysInf.NoMealCmtPM15.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If SysInf.NoMealCmtPM15.Trim <> "" AndAlso blnFlg = False Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd15
                                            If SysInf.NoMealCmtPM15.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If blnFlg = False Then
                                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                    End If
                                End If


                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealL, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMeal15, String)
                        End If

                        '夕
                        If objRecOrderSTD.NoMealE = -1 Then
                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                Dim blnFlg As Boolean = False
                                '欠食コメントがあるかチェック
                                If SysInf.NoMealCmtEve.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If SysInf.NoMealCmtEve.Trim <> "" AndAlso blnFlg = False Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCdE
                                        If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If blnFlg = False Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                End If
                            End If


                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealE, String)
                        End If

                        '20時
                        If objRecOrderSTD.NoMeal20 = -1 Then
                            '夕
                            If objRecOrderSTD.NoMealE = -1 Then
                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                                End If

                                If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.NightDiv) Then
                                    strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.NightDiv), String)
                                End If

                                ''欠食かどうかチェック
                                If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                Else
                                    Dim blnFlg As Boolean = False
                                    '欠食コメントがあるかチェック
                                    If SysInf.NoMealCmtPM20.Trim <> "" Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd
                                            If SysInf.NoMealCmtPM20.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If SysInf.NoMealCmtPM20.Trim <> "" AndAlso blnFlg = False Then
                                        For Each strCmt As String In objRecOrderSTD.CmtCd20
                                            If SysInf.NoMealCmtPM20.IndexOf(strCmt) >= 0 Then
                                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                                blnFlg = True
                                            End If
                                        Next
                                    End If

                                    If blnFlg = False Then
                                        strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                    End If
                                End If

                            Else
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealE, String)
                            End If
                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMeal20, String)
                        End If

                    Else

                        '朝
                        If objRecOrderSTD.NoMealM = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.MorningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.MorningDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                '欠食コメントがあるかチェック
                                Dim blnFlg As Boolean = False
                                If SysInf.NoMealCmtMon.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If SysInf.NoMealCmtMon.Trim <> "" AndAlso blnFlg = False Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCdM
                                        If SysInf.NoMealCmtMon.IndexOf(strCmt) >= 0 Then
                                            strRtn = CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If blnFlg = False Then
                                    strRtn = CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                End If
                            End If


                        Else
                            strRtn = CType(objRecOrderSTD.NoMealM, String)
                        End If

                        '昼
                        If objRecOrderSTD.NoMealL = -1 Then

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.LunchDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.LunchDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                '欠食コメントがあるかチェック
                                Dim blnFlg As Boolean = False
                                If SysInf.NoMealCmtLun.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If SysInf.NoMealCmtLun.Trim <> "" AndAlso blnFlg = False Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCdL
                                        If SysInf.NoMealCmtLun.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If blnFlg = False Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                End If

                            End If


                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealL, String)
                        End If

                        '夕
                        If objRecOrderSTD.NoMealE = -1 Then
                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.OneDay) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.OneDay), String)
                            End If

                            If objHashMeal.ContainsKey(ClsCst.enmEatTimeDiv.EveningDiv) Then
                                strWkMealCd = CType(objHashMeal.Item(ClsCst.enmEatTimeDiv.EveningDiv), String)
                            End If

                            ''欠食かどうかチェック
                            If CheckNoMealCd(SysInf, strWkMealCd) = True Then
                                strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                            Else
                                '欠食コメントがあるかチェック
                                Dim blnFlg As Boolean = False
                                If SysInf.NoMealCmtEve.Trim <> "" Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCd
                                        If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If SysInf.NoMealCmtEve.Trim <> "" AndAlso blnFlg = False Then
                                    For Each strCmt As String In objRecOrderSTD.CmtCdE
                                        If SysInf.NoMealCmtEve.IndexOf(strCmt) >= 0 Then
                                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.NoMeal, String)
                                            blnFlg = True
                                        End If
                                    Next
                                End If

                                If blnFlg = False Then
                                    strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(ClsCst.PerCST.enmNoMeal.Normal, String)
                                End If

                            End If


                        Else
                            strRtn &= ORDER_DELIMITER_CHAR_PIPE & CType(objRecOrderSTD.NoMealE, String)
                        End If

                    End If

                End If

            End If

            Return strRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangComExecDiv
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
    Private Function pfChangComExecDiv(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                       ByVal objExecDivHash As SortedList) As Integer

        Dim intRtn As Integer = 0

        Try

            Dim strKey As String

            strKey = ""
            strKey &= objRecOrderSTD.TelMDiv.Trim
            strKey &= objRecOrderSTD.ExecDiv.Trim

            ''Hashがある場合
            If Not objExecDivHash Is Nothing Then
                ''データがあった場合
                If objExecDivHash.ContainsKey(strKey) Then
                    intRtn = CType(objExecDivHash.Item(strKey), Integer)
                Else
                    intRtn = CType(objRecOrderSTD.ExecDiv, Integer)
                End If
            Else
                intRtn = CType(objRecOrderSTD.ExecDiv, Integer)
            End If

            Return intRtn

        Catch ex As Exception

            Throw

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangComOdDiv
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
    Private Function pfChangComOdDiv(ByVal objRecOrderSTD As Rec_OrderSTD, _
                                     ByVal objOdDivHash As SortedList) As String

        Dim strRtn As String = ""

        Try

            Dim strKey As String

            strKey = ""
            strKey &= objRecOrderSTD.TelMDiv.Trim
            strKey &= objRecOrderSTD.OdDiv.Trim
            strKey &= objRecOrderSTD.ChgRule1.Trim
            strKey &= objRecOrderSTD.ChgRule2.Trim

            ''Hashがある場合
            If Not objOdDivHash Is Nothing Then
                ''データがあった場合
                If objOdDivHash.ContainsKey(strKey) Then
                    strRtn = CType(objOdDivHash.Item(strKey), String)
                Else

                    ''データがない場合
                    strKey = ""
                    strKey &= objRecOrderSTD.TelMDiv.Trim
                    strKey &= objRecOrderSTD.OdDiv.Trim
                    strKey &= objRecOrderSTD.ChgRule1.Trim

                    ''データがあった場合
                    If objOdDivHash.ContainsKey(strKey) Then
                        strRtn = CType(objOdDivHash.Item(strKey), String)
                    Else

                        ''データがない場合
                        strKey = ""
                        strKey &= objRecOrderSTD.TelMDiv.Trim
                        strKey &= objRecOrderSTD.OdDiv.Trim

                        ''データがあった場合
                        If objOdDivHash.ContainsKey(strKey) Then
                            strRtn = CType(objOdDivHash.Item(strKey), String)
                        End If
                    End If
                End If
            End If

            Return strRtn

        Catch ex As Exception
            Throw
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangStYMD
    '
    '----------------------------------------------------------------------
    ' Comment    :
    '
    '----------------------------------------------------------------------
    ' Design On  :2012/05/08
    '        By  :FTX)S.Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Private Function pfChangStYMD(ByVal strStYMD As String, _
                                  ByVal intStTmDiv As Integer, _
                                  ByVal strMoveYMD As String, _
                                  ByVal strMovTime As String, _
                                  ByVal SysInf As Rec_XSys) As String

        Dim strRtn As String

        Try

            strRtn = strStYMD






            Return strRtn

        Catch ex As Exception

            Return strStYMD

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfChangWard
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
    Private Function pfChangWard(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = strValue

            ''Hashがある場合
            If Not m_hashWardCdChg Is Nothing Then
                ''データがあった場合
                If m_hashWardCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashWardCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangRoom
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
    Private Function pfChangRoom(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = strValue

            ''Hashがある場合
            If Not m_hashRoomCdChg Is Nothing Then
                ''データがあった場合
                If m_hashRoomCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashRoomCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangWard
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
    Private Function pfChangMedCare(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = strValue

            ''Hashがある場合
            If Not m_hashCoStCdChg Is Nothing Then
                ''データがあった場合
                If m_hashCoStCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashCoStCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangMealCd
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
    Private Function pfChangMealCd(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashMealCdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashMealCdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashMealCdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfChangMaCkCd
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
    Private Function pfChangMaCkCd(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashMainFdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashMainFdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashMainFdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangTubeNut
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
    Private Function pfChangTubeNut(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashTubeFdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashTubeFdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashTubeFdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangMilk
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
    Private Function pfChangMilk(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashMilkFdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashMilkFdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashMilkFdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangSdShap
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
    Private Function pfChangSdShap(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashSdFmCdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashSdFmCdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashSdFmCdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangDelivCd
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
    Private Function pfChangDelivCd(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashDlvyCdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashDlvyCdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashDlvyCdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangDelivCd
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
    Private Function pfChangIllCd(ByVal strValue As String) As String

        Dim strRtn As String

        Try

            strRtn = strValue

            If strRtn <> "" Then

                ''Hashがある場合
                If Not m_hashSickCdChg Is Nothing Then
                    ''データがあった場合
                    If m_hashSickCdChg.ContainsKey(strValue.Trim) Then
                        strRtn = CType(m_hashSickCdChg.Item(strValue.Trim), String)
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strValue

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangKinCmt
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
    Private Function pfChangKinCmt(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = ""

            ''Hashがある場合
            If Not m_hashCmtKinCdChg Is Nothing Then
                ''データがあった場合
                If m_hashCmtKinCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashCmtKinCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception

            Return ""

        End Try

    End Function


    '**********************************************************************
    ' Name       :pfChangKinCmt
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
    Private Function pfChangCmt(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = ""

            ''Hashがある場合
            If Not m_hashCmtCdChg Is Nothing Then
                ''データがあった場合
                If m_hashCmtCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashCmtCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception

            Return ""

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangTubeMealMethod
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
    Private Function pfChangTubeMealMethod(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = strValue

            ''Hashがある場合
            If Not m_hashTubeMealMethod Is Nothing Then
                ''データがあった場合
                If m_hashTubeMealMethod.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashTubeMealMethod.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception
            Return strValue
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangJiCd
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
    Private Function pfChangJiCd(ByVal strValue As String) As String

        Dim strRtn As String

        Try
            strRtn = strValue

            ''Hashがある場合
            If Not m_hashJiCdChg Is Nothing Then
                ''データがあった場合
                If m_hashJiCdChg.ContainsKey(strValue.Trim) Then
                    strRtn = CType(m_hashJiCdChg.Item(strValue.Trim), String)
                End If
            End If

            Return strRtn

        Catch ex As Exception
            Return strValue
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNewKinCmtCd
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
    Private Function pfGetNewCmtCd(ByVal strOldOrderCmt As String, _
                                   ByVal intEatTmDiv As Integer, _
                                   ByVal objAllCmtAjiHash As Hashtable, _
                                   ByVal objCmtMHash As Hashtable, _
                                   ByVal objBumonCmtMHash As Hashtable, _
                                   ByVal intCmtCls As enumCmtCls, _
                                   ByRef blnBumon As Boolean) As String

        Dim strRtn As String

        Try

            Dim strTrimOldOrderCmt As String

            Dim strChangeCode As String
            Dim objRecMCmt As New Rec_MCmt

            strRtn = ""
            blnBumon = False

            strTrimOldOrderCmt = strOldOrderCmt.Trim

            ''禁止コメントの場合
            If intCmtCls = enumCmtCls.KinCmt Then

                ''変換コード
                strChangeCode = pfChangKinCmt(strTrimOldOrderCmt)

            Else

                ''変換コード
                strChangeCode = pfChangCmt(strTrimOldOrderCmt)

            End If

            If strChangeCode <> "" Then

                ''変換する場合
                If objAllCmtAjiHash.ContainsKey(strChangeCode) Then

                    objRecMCmt = CType(objAllCmtAjiHash.Item(strChangeCode), Rec_MCmt)

                    Select Case intEatTmDiv

                        Case ClsCst.enmEatTimeDiv.MorningDiv

                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.AM10Div

                            If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.LunchDiv

                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.PM15Div

                            If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.EveningDiv

                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.NightDiv

                            If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                    End Select


                End If

            Else

                ''変換しない場合
                ''オーダのマスタにあった場合
                If objCmtMHash.ContainsKey(strTrimOldOrderCmt) Then

                    objRecMCmt = CType(objCmtMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                    Select Case intEatTmDiv

                        Case ClsCst.enmEatTimeDiv.MorningDiv

                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.AM10Div

                            If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.LunchDiv

                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.PM15Div

                            If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.EveningDiv

                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.NightDiv

                            If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                    End Select

                Else

                    ''部門コメント
                    If Not objBumonCmtMHash Is Nothing Then

                        ''部門マスタにあった場合
                        If objBumonCmtMHash.ContainsKey(strTrimOldOrderCmt) Then

                            objRecMCmt = CType(objBumonCmtMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                            Select Case intEatTmDiv

                                Case ClsCst.enmEatTimeDiv.MorningDiv

                                    If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                                Case ClsCst.enmEatTimeDiv.AM10Div

                                    If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                                Case ClsCst.enmEatTimeDiv.LunchDiv

                                    If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                                Case ClsCst.enmEatTimeDiv.PM15Div

                                    If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                                Case ClsCst.enmEatTimeDiv.EveningDiv

                                    If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                                Case ClsCst.enmEatTimeDiv.NightDiv

                                    If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                        blnBumon = True
                                        strRtn = objRecMCmt.CmtCd.Trim
                                    End If

                            End Select

                        End If

                    End If

                End If

                ''マスタにない場合
                If objRecMCmt Is Nothing OrElse objRecMCmt.CmtCd = "" Then
                    If strRtn.Trim = "" Then
                        strRtn = strTrimOldOrderCmt
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strOldOrderCmt

        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetNewKinCmtCd
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
    Private Function pfGetNewCmtCd(ByVal strOldOrderCmt As String, _
                                   ByVal intEatTmDiv As Integer, _
                                   ByVal objAllCmtAjiHash As Hashtable, _
                                   ByVal objCmtKinMHash As Hashtable, _
                                   ByVal objCmtMHash As Hashtable, _
                                   ByVal objBumonCmtKinMHash As Hashtable, _
                                   ByVal objBumonCmtMHash As Hashtable, _
                                   ByRef blnBumon As Boolean) As String

        Dim strRtn As String

        Try

            Dim strTrimOldOrderCmt As String
            Dim strChangeCode As String
            Dim objRecMCmt As New Rec_MCmt
            Dim blnNoCmtInMCmt As Boolean = True

            strRtn = ""
            blnBumon = False

            strTrimOldOrderCmt = strOldOrderCmt.Trim

            ''変換コード
            strChangeCode = pfChangCmt(strTrimOldOrderCmt)

            If strChangeCode <> "" Then

                ''変換する場合
                If objAllCmtAjiHash.ContainsKey(strChangeCode) Then

                    objRecMCmt = CType(objAllCmtAjiHash.Item(strChangeCode), Rec_MCmt)

                    Select Case intEatTmDiv

                        Case ClsCst.enmEatTimeDiv.MorningDiv

                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.AM10Div

                            If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.LunchDiv

                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.PM15Div

                            If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.EveningDiv

                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.NightDiv

                            If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                    End Select


                End If

            Else

                ''変換しない場合
                ''オーダのマスタにあった場合
                If objCmtKinMHash.ContainsKey(strTrimOldOrderCmt) Then

                    objRecMCmt = CType(objCmtKinMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                    Select Case intEatTmDiv

                        Case ClsCst.enmEatTimeDiv.MorningDiv

                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.AM10Div

                            If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.LunchDiv

                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.PM15Div

                            If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.EveningDiv

                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.NightDiv

                            If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                    End Select

                End If

                ''オーダのマスタにあった場合
                If objCmtMHash.ContainsKey(strTrimOldOrderCmt) Then

                    objRecMCmt = CType(objCmtMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                    Select Case intEatTmDiv

                        Case ClsCst.enmEatTimeDiv.MorningDiv

                            If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.AM10Div

                            If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.LunchDiv

                            If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.PM15Div

                            If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.EveningDiv

                            If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                        Case ClsCst.enmEatTimeDiv.NightDiv

                            If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                strRtn = objRecMCmt.CmtCd.Trim
                            End If

                    End Select

                End If

                '部門禁止コメント
                If Not objBumonCmtKinMHash Is Nothing Then

                    ''部門マスタにあった場合
                    If objBumonCmtKinMHash.ContainsKey(strTrimOldOrderCmt) Then

                        objRecMCmt = CType(objBumonCmtKinMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                        Select Case intEatTmDiv

                            Case ClsCst.enmEatTimeDiv.MorningDiv

                                If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.AM10Div

                                If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.LunchDiv

                                If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.PM15Div

                                If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.EveningDiv

                                If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.NightDiv

                                If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                        End Select

                    End If

                End If

                ''部門コメント
                If Not objBumonCmtMHash Is Nothing Then

                    ''部門マスタにあった場合
                    If objBumonCmtMHash.ContainsKey(strTrimOldOrderCmt) Then

                        objRecMCmt = CType(objBumonCmtMHash.Item(strTrimOldOrderCmt), Rec_MCmt)

                        Select Case intEatTmDiv

                            Case ClsCst.enmEatTimeDiv.MorningDiv

                                If objRecMCmt.TgtMorn = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.AM10Div

                                If objRecMCmt.TgtMornK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.LunchDiv

                                If objRecMCmt.TgtLun = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.PM15Div

                                If objRecMCmt.TgtLunK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.EveningDiv

                                If objRecMCmt.TgtEvo = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                            Case ClsCst.enmEatTimeDiv.NightDiv

                                If objRecMCmt.TgtEvoK = ClsCst.enmValidF.Exists Then
                                    blnBumon = True
                                    strRtn = objRecMCmt.CmtCd.Trim
                                End If

                        End Select

                    End If

                End If

                ''マスタにない場合
                If objRecMCmt Is Nothing OrElse objRecMCmt.CmtCd = "" Then
                    If strRtn.Trim = "" Then
                        strRtn = strTrimOldOrderCmt
                    End If
                End If

            End If

            Return strRtn

        Catch ex As Exception

            Return strOldOrderCmt

        End Try

    End Function

    '**********************************************************************
    ' Name       :CheckNoMealCd
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
    Private Function CheckNoMealCd(ByVal SysInf As Rec_XSys, _
                                   ByVal strMealCd As String) As Boolean

        Dim blnRtn As Boolean

        blnRtn = False

        Try

            ''欠食をチェック
            Dim objNoMealCd As New StringCollection

            objNoMealCd = SplitOrderData(SysInf.OrderNoMealCd, ORDER_DELIMITER_CHAR_PIPE)

            If Not objNoMealCd Is Nothing AndAlso objNoMealCd.Count > 0 Then

                For Each strNoMealCd As String In objNoMealCd

                    If strMealCd = strNoMealCd Then

                        blnRtn = True
                        Exit For
                    End If

                Next

            End If

            ''外泊をチェック
            If blnRtn = False Then

                Dim objGoOutMealCd As New StringCollection

                objGoOutMealCd = SplitOrderData(SysInf.OrderGoOutMealCd, ORDER_DELIMITER_CHAR_PIPE)

                If Not objNoMealCd Is Nothing AndAlso objNoMealCd.Count > 0 Then

                    For Each strNoMealCd As String In objNoMealCd

                        If strMealCd = strNoMealCd Then

                            blnRtn = True
                            Exit For
                        End If

                    Next

                End If

            End If

            Return blnRtn

        Catch ex As Exception
            Return blnRtn
        End Try

    End Function

    '**********************************************************************
    ' Name       :CheckNoMainFdCd
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
    Private Function CheckNoMainFdCd(ByVal SysInf As Rec_XSys, _
                                     ByVal strOrderMainFdCd As String) As Boolean

        Dim blnRtn As Boolean

        blnRtn = False

        Try

            If strOrderMainFdCd <> "" Then

                ''欠食の主食をチェック
                Dim objNoMainFdCd As New StringCollection

                objNoMainFdCd = SplitOrderData(SysInf.OrderNoMainFdCd, ORDER_DELIMITER_CHAR_PIPE)

                If Not objNoMainFdCd Is Nothing AndAlso objNoMainFdCd.Count > 0 Then

                    For Each strNoMainFdCd As String In objNoMainFdCd

                        If strOrderMainFdCd = strNoMainFdCd Then

                            blnRtn = True
                            Exit For
                        End If

                    Next

                End If

            End If

            Return blnRtn

        Catch ex As Exception
            Return blnRtn
        End Try

    End Function

    '**********************************************************************
    ' Name       :CheckBreadFdCd
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
    Private Function CheckBreadFdCd(ByVal SysInf As Rec_XSys, _
                                    ByVal strOrderMainFdCd As String) As Boolean

        Dim blnRtn As Boolean

        blnRtn = False

        Try

            If strOrderMainFdCd <> "" Then

                ''パンの主食をチェック
                Dim objBreadFdCd As New StringCollection

                objBreadFdCd = SplitOrderData(SysInf.OrderBreadFdCd, ORDER_DELIMITER_CHAR_PIPE)

                If Not objBreadFdCd Is Nothing AndAlso objBreadFdCd.Count > 0 Then

                    For Each strBreadFdCd As String In objBreadFdCd

                        If strOrderMainFdCd = strBreadFdCd Then

                            blnRtn = True
                            Exit For
                        End If

                    Next

                End If

            End If

            Return blnRtn

        Catch ex As Exception
            Return blnRtn
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfChangRoom
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
    Private Function GetNoMealRs(ByVal SysInf As Rec_XSys, _
                                 ByVal strMealCd As String) As String

        Dim strRtn As String

        strRtn = ""

        Try

            ''欠食をチェック
            Dim objNoMealCd As New StringCollection

            objNoMealCd = SplitOrderData(SysInf.OrderNoMealCd, ORDER_DELIMITER_CHAR_PIPE)

            If Not objNoMealCd Is Nothing AndAlso objNoMealCd.Count > 0 Then

                For Each strNoMealCd As String In objNoMealCd

                    If strMealCd = strNoMealCd Then

                        strRtn = SysInf.OrderNoMlRs
                        Exit For
                    End If

                Next

            End If

            ' ''外泊をチェック
            'If strRtn = "" Then

            '    Dim objGoOutMealCd As New StringCollection

            '    objGoOutMealCd = SplitOrderData(SysInf.OrderGoOutMealCd, ORDER_DELIMITER_CHAR_PIPE)

            '    If Not objNoMealCd Is Nothing AndAlso objNoMealCd.Count > 0 Then

            '        For Each strNoMealCd As String In objNoMealCd

            '            If strMealCd = strNoMealCd Then

            '                strRtn = SysInf.OrderNoMlRs
            '                Exit For
            '            End If

            '        Next

            '    End If

            'End If

            Return strRtn

        Catch ex As Exception
            Return strRtn
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetFirstStringFromCollection
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
    Private Function pfGetFirstStringFromCollection(ByVal strColValue As String, _
                                                    ByVal strDefaultValue As String) As String

        Dim strRtn As String

        strRtn = strDefaultValue

        Try

            Dim objStrCol As New StringCollection

            objStrCol = SplitOrderData(strColValue, ORDER_DELIMITER_CHAR_PIPE)

            ''値があれば
            If Not objStrCol Is Nothing AndAlso objStrCol.Count > 0 Then
                strRtn = objStrCol.Item(0)
            End If

            Return strRtn

        Catch ex As Exception
            Return strRtn
        End Try

    End Function

    '**********************************************************************
    ' Name       :pfGetStringFromCollection
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
    Private Function pfGetIndexStringFromCollection(ByVal strColValue As String, _
                                                    ByVal intIndex As Integer, _
                                                    ByVal strDefaultValue As String) As String

        Dim strRtn As String

        strRtn = strDefaultValue

        Try

            Dim objStrCol As New StringCollection

            objStrCol = SplitOrderData(strColValue, ORDER_DELIMITER_CHAR_PIPE)

            ''値があれば
            If Not objStrCol Is Nothing AndAlso objStrCol.Count > 0 Then

                If intIndex <= objStrCol.Count Then
                    strRtn = objStrCol.Item(intIndex)
                Else
                    strRtn = strDefaultValue
                End If

            End If

            Return strRtn

        Catch ex As Exception
            Return strRtn
        End Try

    End Function


    '**********************************************************************
    ' Name       :pfGetStringFromCollection
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
    Private Function pfGetIndexCountFromCollection(ByVal strColValue As String) As Integer

        Dim intRtn As Integer

        Try

            Dim objStrCol As New StringCollection

            objStrCol = SplitOrderData(strColValue, ORDER_DELIMITER_CHAR_PIPE)

            ''値があれば
            If Not objStrCol Is Nothing AndAlso objStrCol.Count > 0 Then
                intRtn = objStrCol.Count
            Else
                intRtn = 0
            End If

            Return intRtn

        Catch ex As Exception
            Return -1
        End Try

    End Function

    '**********************************************************************
    ' Name       :SetHashCdChg
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
    Public Function SetHashMstCdChg(ByVal objColOMCdChg As Col_OMCdChg, _
                                    ByVal objColMNameChg As Col_MName(), _
                                    ByVal objMealMHash As Hashtable, _
                                    ByVal objColMCmt As Col_MCmt, _
                                    ByVal objColMCok As Col_MCok, _
                                    ByVal SysInf As Rec_XSys) As Boolean

        Try

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／　　　　　　　　　　　　　　　　　　　　　　　　　　　　　＼
            ''コード変換テーブルのハッシュを作成
            For Each objRecOMCdChg As Rec_OMCdChg In objColOMCdChg

                Select Case objRecOMCdChg.TblNm.Trim

                    Case ClsCst.TableName.M_Name

                        '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                        '／名称マスタの場合　　　　　　　　　　　　　　　　　　　　　＼
                        Select Case objRecOMCdChg.DisTp.Trim

                            Case ClsCst.MNameDisCd.Ward             '病棟

                                ''コードがない場合
                                If m_hashWardCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashWardCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If


                            Case ClsCst.MNameDisCd.SickRoom         '病室

                                ''コードがない場合
                                If m_hashRoomCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashRoomCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case ClsCst.MNameDisCd.Consulting       '診療科

                                ''コードがない場合
                                If m_hashCoStCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashCoStCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case ClsCst.MNameDisCd.SickName         '病名

                                ''コードがない場合
                                If m_hashSickCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashSickCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case ClsCst.MNameDisCd.SddishesForm     '副食形態

                                ''コードがない場合
                                If m_hashSdFmCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashSdFmCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case ClsCst.MNameDisCd.DeliverPlace     '配膳先情報

                                ''コードがない場合
                                If m_hashDlvyCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashDlvyCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case ClsCst.MNameDisCd.NutMealMethod    '給与方法

                                ''コードがない場合
                                If m_hashTubeMealMethod.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashTubeMealMethod.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If


                        End Select
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    Case ClsCst.TableName.M_Cmt

                        '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                        '／コメントマスタの場合　　　　　　　　　　　　　　　　　　　＼
                        If objRecOMCdChg.DisTp.Trim = "0" OrElse objRecOMCdChg.DisTp.Trim = "" Then

                            ''通常コメントの場合

                            ''コードがない場合
                            If m_hashCmtCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                m_hashCmtCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                            End If

                        Else

                            ''禁止コメントの場合

                            ''コードがない場合
                            If m_hashCmtKinCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                m_hashCmtKinCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                            End If

                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    Case ClsCst.TableName.M_MealTpP

                        '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                        '／食種マスタ（個人）の場合　　　　　　　　　　　　　　　　　＼
                        If m_hashMealCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                            m_hashMealCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    Case ClsCst.TableName.M_Cok

                        '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                        '／料理マスタの場合　　　　　　　　　　　　　　　　　　　　　＼
                        Select Case objRecOMCdChg.DisTp.Trim

                            Case CType(ClsCst.enmCkDiv.MainCok, String)    '主食

                                ''コードがない場合
                                If m_hashMainFdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashMainFdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case CType(ClsCst.enmCkDiv.Tube, String)      '経管食

                                ''コードがない場合
                                If m_hashTubeFdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashTubeFdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                            Case CType(ClsCst.enmCkDiv.Milk, String)      '調乳

                                ''コードがない場合
                                If m_hashMilkFdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                                    m_hashMilkFdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                                End If

                        End Select
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                    Case ClsCst.TableName.M_JiGyo

                        '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                        '／事業所マスタの場合　　　　　　　　　　　　　　　　　＼
                        If m_hashJiCdChg.ContainsKey(objRecOMCdChg.OrderCd.Trim) = False Then
                            m_hashJiCdChg.Add(objRecOMCdChg.OrderCd.Trim, objRecOMCdChg.AjiDASCd.Trim)
                        End If
                        '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                End Select

            Next
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／　　　　　　　　　　　　　　　　　　　　　　　　　　　　　＼
            ''名称マスタのハッシュを作成
            If Not objColMNameChg Is Nothing AndAlso objColMNameChg.Count > 0 Then

                For Each objColMName As Col_MName In objColMNameChg

                    If Not objColMName Is Nothing AndAlso objColMName.Count > 0 Then

                        For Each objRecMName As Rec_MName In objColMName

                            With objRecMName

                                ''オーダコードが空白ではない場合
                                If .OrdCd.Trim <> "" Then

                                    '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                                    '／名称マスタの場合　　　　　　　　　　　　　　　　　　　　　＼
                                    Select Case .DisCd

                                        Case ClsCst.MNameDisCd.Ward             '病棟

                                            ''コードがない場合
                                            If m_hashWardCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashWardCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                            ''病棟による食種変換がある場合
                                            If SysInf.OrderChangeCd.IndexOf(ClsCst.ORecv.OCHG_OrderWardCnvSet) <> -1 Then

                                                If m_hashWardMealAddCd.ContainsKey(.OrdCd.Trim) = False Then
                                                    m_hashWardMealAddCd.Add(.OrdCd.Trim, .CdVal3.Trim)
                                                End If

                                            End If

                                        Case ClsCst.MNameDisCd.SickRoom         '病室

                                            ''コードがない場合
                                            If m_hashRoomCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashRoomCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                        Case ClsCst.MNameDisCd.Consulting       '診療科

                                            ''コードがない場合
                                            If m_hashCoStCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashCoStCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                        Case ClsCst.MNameDisCd.SickName         '病名

                                            ''コードがない場合
                                            If m_hashSickCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashSickCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                        Case ClsCst.MNameDisCd.SddishesForm     '副食形態

                                            ''コードがない場合
                                            If m_hashSdFmCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashSdFmCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                        Case ClsCst.MNameDisCd.DeliverPlace     '配膳先情報

                                            ''コードがない場合
                                            If m_hashDlvyCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashDlvyCdChg.Add(.OrdCd.Trim, .Cd1.Trim)
                                            End If

                                        Case ClsCst.MNameDisCd.NutMealMethod     '給与方法

                                            ''コードがない場合は「名前」を追加する
                                            If m_hashTubeMealMethod.ContainsKey(.OrdCd.Trim) = False Then
                                                m_hashTubeMealMethod.Add(.OrdCd.Trim, .CdNm.Trim)
                                            End If

                                    End Select
                                    '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                                End If

                            End With

                        Next

                    End If

                Next

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／　　　　　　　　　　　　　　　　　　　　　　　　　　　　　＼
            ''食種マスタのハッシュを作成
            Dim objMMealTpP As Rec_MMealTpP

            If Not objMealMHash Is Nothing AndAlso objMealMHash.Count > 0 Then

                For Each decMMealTpP As DictionaryEntry In objMealMHash

                    ''更新する画面の情報を取得
                    objMMealTpP = CType(decMMealTpP.Value, Rec_MMealTpP)

                    With objMMealTpP

                        If .OrdCd.Trim <> "" Then

                            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                            '／食種マスタ（個人）の場合　　　　　　　　　　　　　　　　　＼
                            If m_hashMealCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                m_hashMealCdChg.Add(.OrdCd.Trim, .MealCd.Trim)
                            End If
                            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                        End If

                    End With

                Next

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／　　　　　　　　　　　　　　　　　　　　　　　　　　　　　＼
            ''コメントマスタのハッシュを作成
            If Not objColMCmt Is Nothing AndAlso objColMCmt.Count > 0 Then

                For Each objMCmt As Rec_MCmt In objColMCmt

                    With objMCmt

                        If .OrdCd.Trim <> "" Then

                            If .CmtCls = 0 Then     '通常

                                ''通常コメントの場合

                                ''コードがない場合
                                If m_hashCmtCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                    m_hashCmtCdChg.Add(.OrdCd.Trim, .CmtCd.Trim)
                                End If

                            Else

                                ''禁止コメントの場合

                                ''コードがない場合
                                If m_hashCmtKinCdChg.ContainsKey(.OrdCd.Trim) = False Then
                                    m_hashCmtKinCdChg.Add(.OrdCd.Trim, .CmtCd.Trim)
                                End If

                            End If

                        End If

                    End With

                Next

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            '　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
            '／　　　　　　　　　　　　　　　　　　　　　　　　　　　　　＼
            ''料理マスタのハッシュを作成
            If Not objColMCok Is Nothing AndAlso objColMCok.Count > 0 Then

                For Each objMCok As Rec_MCok In objColMCok

                    With objMCok

                        If .OrdCd.Trim <> "" Then

                            Select Case .CkDiv

                                Case ClsCst.enmCkDiv.MainCok    '主食

                                    ''コードがない場合
                                    If m_hashMainFdChg.ContainsKey(.OrdCd.Trim) = False Then
                                        m_hashMainFdChg.Add(.OrdCd.Trim, .CkCd.Trim)
                                    End If

                                Case ClsCst.enmCkDiv.Tube      '経管食

                                    ''コードがない場合
                                    If m_hashTubeFdChg.ContainsKey(.OrdCd.Trim) = False Then
                                        m_hashTubeFdChg.Add(.OrdCd.Trim, .CkCd.Trim)
                                    End If

                                Case ClsCst.enmCkDiv.Milk      '経管食

                                    ''コードがない場合
                                    If m_hashMilkFdChg.ContainsKey(.OrdCd.Trim) = False Then
                                        m_hashMilkFdChg.Add(.OrdCd.Trim, .CkCd.Trim)
                                    End If


                            End Select

                        End If

                    End With

                Next

            End If
            '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function

    '**********************************************************************
    ' Summary    : 
    '----------------------------------------------------------------------
    ' Method     : Public Function WriteOrderDataToFile()
    '
    ' Parameter  : 
    '
    ' Return     :
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2011/7/12
    '        By  : FTC) S_Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function WriteOrderDataToFile(ByVal strFilePath As String, _
                                         ByVal objCOLORecv As Col_ORecv) As Boolean

        Try

            Dim strLogFilePathName As String
            Dim strLogString As String

            If Not objCOLORecv Is Nothing AndAlso objCOLORecv.Count > 0 Then

                ''書き込むファイルパス
                strLogFilePathName = strFilePath

                ''　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                ''／ファイルへの書き出しを行う処理　　　　　　　　　　　　　　＼
                Dim objFileStream As New IO.FileStream(strLogFilePathName, _
                  IO.FileMode.Append, IO.FileAccess.Write, IO.FileShare.Write)
                Dim objTextWriterTraceListener As New TextWriterTraceListener(objFileStream)

                For Each objReORecv As Rec_ORecv In objCOLORecv

                    With objReORecv

                        strLogString = ""
                        strLogString &= .OpYMD

                        strLogString &= vbTab & .OpNo
                        strLogString &= vbTab & .OdTp
                        strLogString &= vbTab & .RfctDiv
                        strLogString &= vbTab & .PpIsDiv
                        strLogString &= vbTab & .OdNo
                        strLogString &= vbTab & .Rsheet
                        strLogString &= vbTab & .AtvOd
                        strLogString &= vbTab & .Deadline
                        strLogString &= vbTab & .ExecDiv
                        strLogString &= vbTab & .TelMDiv
                        strLogString &= vbTab & .OdDiv
                        strLogString &= vbTab & .JiCd
                        strLogString &= vbTab & .PerCd
                        strLogString &= vbTab & .StYMD
                        strLogString &= vbTab & .StTmDiv
                        strLogString &= vbTab & .EdYMD
                        strLogString &= vbTab & .EdTmDiv
                        strLogString &= vbTab & .MovYMD
                        strLogString &= vbTab & .MovTime
                        strLogString &= vbTab & .OdData
                        strLogString &= vbTab & .RefCol
                        strLogString &= vbTab & .RefCol2
                        strLogString &= vbTab & .RefCol3
                        strLogString &= vbTab & .ErrTp
                        strLogString &= vbTab & .ErrNo
                        strLogString &= vbTab & .ErrMsg
                        strLogString &= vbTab & Now.ToString("yyyy/MM/dd")
                        strLogString &= vbTab & .UpDtUser

                    End With

                    ''ファイル書込
                    Trace.Listeners.Add(objTextWriterTraceListener)

                    objTextWriterTraceListener.WriteLine(strLogString)

                Next

                objTextWriterTraceListener.Flush()
                objTextWriterTraceListener.Close()
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                'メモリリーク対応
                Trace.Listeners.Clear()
                objFileStream.Close()

            End If

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function


    '**********************************************************************
    ' Summary    : 
    '----------------------------------------------------------------------
    ' Method     : Public Function WriteStringColDataToFile()
    '
    ' Parameter  : 
    '
    ' Return     :
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2011/7/12
    '        By  : FTC) S_Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function WriteStringColDataToFile(ByVal strFilePath As String, _
                                             ByVal objStrCol As StringCollection) As Boolean

        Try

            Dim strLogFilePathName As String
            Dim strLogString As String

            If Not objStrCol Is Nothing AndAlso objStrCol.Count > 0 Then

                ''書き込むファイルパス
                strLogFilePathName = strFilePath

                ''　＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿
                ''／ファイルへの書き出しを行う処理　　　　　　　　　　　　　　＼
                Dim objFileStream As New IO.FileStream(strLogFilePathName, _
                  IO.FileMode.Append, IO.FileAccess.Write, IO.FileShare.Write)
                Dim objTextWriterTraceListener As New TextWriterTraceListener(objFileStream)

                For Each strRec As String In objStrCol

                    strLogString = strRec

                    ''ファイル書込
                    Trace.Listeners.Add(objTextWriterTraceListener)

                    objTextWriterTraceListener.WriteLine(strLogString)

                Next

                objTextWriterTraceListener.Flush()
                objTextWriterTraceListener.Close()
                '＼＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿／

                'メモリリーク対応
                Trace.Listeners.Clear()
                objFileStream.Close()

            End If

            Return True

        Catch ex As Exception

            Return False

        End Try

    End Function

    '**********************************************************************
    ' Summary    : 
    '----------------------------------------------------------------------
    ' Method     : Public Sub CreateBackUpOrderFileFolder()
    '
    ' Parameter  : 
    '
    ' Return     :
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2011/7/12
    '        By  : FTC) S_Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function CreateOrderFileBackUpFolder(ByVal strBackUpFilePath As String, _
                                                ByVal strEntryDateYMD As String) As String

        Dim strRtn As String = ""

        Try

            Dim strNewFilePath As String = ""

            ''新しいパスを作成
            strNewFilePath = IO.Path.Combine(strEntryDateYMD.Substring(0, 4), strEntryDateYMD.Substring(4, 2))
            strNewFilePath = IO.Path.Combine(strNewFilePath, strEntryDateYMD.Substring(6, 2))

            If strBackUpFilePath Is Nothing OrElse strBackUpFilePath.Trim = "" Then
                strBackUpFilePath = "C:\AjiDAS7\Order"
            End If

            strBackUpFilePath = IO.Path.Combine(strBackUpFilePath, strNewFilePath)

            ''年月日のバックアップ用ディレクトリーを作成
            If IO.Directory.Exists(strBackUpFilePath) = False Then
                IO.Directory.CreateDirectory(strBackUpFilePath)
            End If

            strRtn = strBackUpFilePath

            Return strRtn

        Catch ex As Exception

            Return strBackUpFilePath

        End Try

    End Function

    '**********************************************************************
    ' Summary    : 
    '----------------------------------------------------------------------
    ' Method     : Public Sub CreateBackUpOrderFileFolder()
    '
    ' Parameter  : 
    '
    ' Return     :
    '
    '----------------------------------------------------------------------
    ' Comment    : 
    '
    '----------------------------------------------------------------------
    ' Design On  : 2011/7/12
    '        By  : FTC) S_Shin
    '----------------------------------------------------------------------
    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Function CreateOrderBackUpFileName(ByVal strBackUpFileName As String, _
                                              ByVal strEntryDateYMDHMS As String) As String

        Dim strRtn As String = ""

        Try
            If strBackUpFileName Is Nothing OrElse strBackUpFileName.Trim = "" Then
                strBackUpFileName = "AjiDAS7Recv.log"
            End If

            ''受信ログファイル名
            strRtn = strEntryDateYMDHMS & strBackUpFileName

            Return strRtn

        Catch ex As Exception
            Return strBackUpFileName
        End Try

    End Function

    '**********************************************************************
    ' Name       :WardCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :WardCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property WardCdChg() As Hashtable
        Get
            Return m_hashWardCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashWardCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :RoomCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :RoomCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property RoomCdChg() As Hashtable
        Get
            Return m_hashRoomCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashRoomCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :CoStCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :CoStCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property CoStCdChg() As Hashtable
        Get
            Return m_hashCoStCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashCoStCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :SickCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :SickCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property SickCdChg() As Hashtable
        Get
            Return m_hashSickCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashSickCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :SickCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :SickCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property SdFmCdChg() As Hashtable
        Get
            Return m_hashSdFmCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashSdFmCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :MealCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :MealCdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property MealCdChg() As Hashtable
        Get
            Return m_hashMealCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashMealCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :MainFdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :MainFdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property MainFdChg() As Hashtable
        Get
            Return m_hashMainFdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashMainFdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :CmtCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :MainFdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property CmtCdChg() As Hashtable
        Get
            Return m_hashCmtCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashCmtCdChg = Value
        End Set
    End Property

    '**********************************************************************
    ' Name       :CmtKinCdChg
    '
    '----------------------------------------------------------------------
    ' Comment    :MainFdChg
    '
    '----------------------------------------------------------------------
    ' Design On  :2011/11/08
    '        By  :FTC)S.Shin
    '----------------------------------------------------------------------

    ' Update On  :
    '        By  :
    ' Comment    :
    '**********************************************************************
    Public Property CmtKinCdChg() As Hashtable
        Get
            Return m_hashCmtKinCdChg
        End Get
        Set(ByVal Value As Hashtable)
            m_hashCmtKinCdChg = Value
        End Set
    End Property

End Class
