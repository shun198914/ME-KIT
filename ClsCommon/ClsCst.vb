Option Strict On
'**************************************************************************
' Name       :ClsCst 共通定数
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
'Imports
'**************************************************************************
Public Class ClsCst
    '**********************************************************************
    'コンスタント定数
    '**********************************************************************
    '**********************************************************************
    'フォルダ名
    '**********************************************************************
    Public Class FldrName
        Public Const AppRoot As String = "%AppRoot%"               'アプリケーション全体のメインフォルダ象徴文字(FrmMenu.exeのあるパス)
        Public Const Image As String = "Image"                     '画像の入っているフォルダ
        Public Const Sound As String = "Sound"                     'サウンドの入っているフォルダ
        Public Const Bin As String = "Bin"                         '実行ファイルの入っているフォルダ
        Public Const Log As String = "Log"                         'ログファイルの入っているフォルダ
        Public Const Order As String = "Order"                     'オーダ連携用ファイルの入っているフォルダ
        Public Const SqlScript As String = "SqlScript"             'SQLスクリプトファイルの入っているフォルダ
        Public Const Help As String = "Help"                       'ヘルプファイルの入っているフォルダ
        Public Const Receive As String = "Receive"                 'ダウンロードするファイル基本フォルダ(以下のサブフォルダの親フォルダ)
        Public Const ReceiveMst As String = Receive & "\Mst"       'マスタデータをダウンロードするフォルダ
        Public Const ReceiveApl As String = Receive & "\Apl"       'アプリケーションをダウンロードするフォルダ
        Public Const ReceiveVpf As String = Receive & "\Vpf"       'ウィルスパターンファイルをダウンロードするフォルダ
        Public Const TempExp As String = "TempExp"                 '圧縮ファイルを解凍するテンポラリフォルダ
        Public Const TempExpMst As String = TempExp & "\Mst"       '圧縮マスタファイルを解凍するテンポラリフォルダ
        Public Const TempExpApl As String = TempExp & "\Apl"       '圧縮アプリケーションを解凍するテンポラリフォルダ
        Public Const TempExpVpf As String = TempExp & "\Vpf"       '圧縮ウィルスパターンファイルを解凍するテンポラリフォルダ
        Public Const ZipBack As String = "ZipBack"                 '圧縮ファイルを保管するフォルダ
        Public Const ZipBackMst As String = ZipBack & "\Mst"       '圧縮マスタファイルを保管するフォルダ
        Public Const ZipBackApl As String = ZipBack & "\Apl"       '圧縮アプリケーションを保管するフォルダ
        Public Const ZipBackVpf As String = ZipBack & "\Vpf"       '圧縮ウィルスパターンファイルを保管するフォルダ
    End Class

    '**********************************************************************
    'ファイル名
    '**********************************************************************
    Public Class FileName
        Public Const ConfigClient As String = "ConfigC.xml"        'XMLクライアント設定ファイル
        Public Const ConfigBackUp As String = "ConfigM.xml"        'XMLバックアップ設定ファイル
        Public Const ConfigServer As String = "ConfigS.xml"        'XMLサーバー設定ファイル
        Public Const LogClient As String = "LogC.txt"              'クライアントログ
        Public Const LogServer As String = "LogS.txt"              'サーバーログ
        Public Const LogAct As String = "ActLog.txt"               '稼動ログ
    End Class

    '**********************************************************************
    '予約画面名
    '**********************************************************************
    Public Class FormName
        Public Const FrmMenu As String = "FrmMenu"                  '常に表示されるメニュー画面(かつFrmMenu.exeのプロセス名)
        Public Const FrmLogIn As String = "FrmLogIn"                'ログイン画面
        Public Const FrmMain As String = "FrmMain"                  'バックグランド画面
        Public Const FrmStartLogo As String = "FrmStartLogo"        'スタートロゴ画面
        Public Const FrmMstFd As String = "FrmMstFd"                '食品マスタ
        Public Const FrmMstFdR As String = "FrmMstFdR"              '食品マスタ連票
        Public Const FrmMstCok As String = "FrmMstCok"              '料理マスタ
        Public Const FrmMstCokR As String = "FrmMstCokR"            '料理マスタ連票
        Public Const FrmFoodSearch As String = "FrmFoodSearch"      '食品検索
        Public Const FrmCookSearch As String = "FrmCookSearch"      '料理検索
        Public Const FrmHaseiTen As String = "FrmHaseiTen"          '発生点画面（材料集計）
        Public Const FrmZaiOut As String = "FrmZaiOut"              '在庫払出画面（１日単位）
        Public Const FrmZaiOutTmDiv As String = "FrmZaiOutTmDiv"        '在庫払出画面（時間区分）
        Public Const FrmPerChangeHospi As String = "FrmPerChangeHospi"  '入院日変更画面
        Public Const FrmZaikoNoDiv As String = "FrmZaikoNoDiv"          '棚卸在庫画面(区分なし)
        Public Const FrmChgJiCd As String = "FrmChgJiCd"                    '事業所コード変更画面
        Public Const FrmClearTable As String = "FrmClearTable"            '新規事業所用データ一括削除画面
        '発注系画面名
        Public Shared HachuForms As String() = New String() {"FrmDlveChange", "FrmHachuEdit", "FrmHachuFood", "FrmHachuHistory" _
                                                           , "FrmHachuList", "FrmHachuSend", "FrmHachuSendOK", "FrmMaterials"}
        '2015/10/06 hirabayashi アイコーメディカル 発注システム改修
        Public Const FrmMTraderMD As String = "FrmMTraderMD"            '納品曜日設定画面
    End Class

    '**********************************************************************
    'ＸＭＬセクション名
    '**********************************************************************
    Public Class xmlSectionName
        Public Const DBSection As String = "DBSection"             'DBセクション

        '*** ↓ *** システムにて未使用の為、コメントアウト Update by K.Aoki *******
        'Public Const ExeToExeRmt As String = "exetoexe.remoting"   'Exe間通信用セクション
        '*** ↑ *** システムにて未使用の為、コメントアウト Update by K.Aoki *******
    End Class

    '**********************************************************************
    'ＸＭＬセクションタグ名
    '**********************************************************************
    Public Class xmlTagName
        Public Const InitialCatalog As String = "InitialCatalog"   'イニシャルカタログ
        Public Const DataSource As String = "DataSource"           'データソース
        Public Const User As String = "User"                       'ユーザ
        Public Const Password As String = "Password"               'パスワード
        Public Const ConnectTimeout As String = "ConnectTimeout"   '接続タイムアウト
        Public Const CommandTimeout As String = "CommandTimeout"   'コマンドタイムアウト
        Public Const Pooling As String = "Pooling"                 '接続プール設定
    End Class

    '**********************************************************************
    'データベースオブジェクト名
    '**********************************************************************
    Public Class DataObjName
        Public Const Table As String = "Table"
    End Class

    '**********************************************************************
    '権限グループ（特権グループのみ）
    '**********************************************************************
    Public Class Kgn
        Public Const System As String = "System"                   'システム開発者
    End Class

    '**********************************************************************
    '長さ
    '**********************************************************************
    Public Class Length
        Public Const CkNm As Integer = 30                          '料理名
    End Class

    '**********************************************************************
    'ＥＮＵＭ定数
    '**********************************************************************
    'ログタイプ
    Public Enum enmLogType
        Information = 1
        Warning = 2
        [Error] = 3
        MethodST = 4
        MethodED = 5
        MsgBox = 6
    End Enum

    'メッセージボックスアイコン
    Public Enum enmMsgBoxIcon
        None = 0
        Information = 1
        Warning = 2
        [Error] = 3
        Question = 4
    End Enum

    '**********************************************************************
    'ＥＮＵＭ定数
    '**********************************************************************
    '個人ログタイプ
    Public Enum enmPerLogType
        SelectGet = 10
        Update = 20
        Insert = 30
        Delete = 40
        Exec = 50
        Print = 60
    End Enum

    Public Const PerLogDelimiter As String = ","

    '**************************************************************************
    'テーブルのコンスト値
    '**************************************************************************
    Public Class TableConst

        Public Const TBL_FdADiv_Normal As Integer = 0           ''指定食品A区分 0=一般
        Public Const TBL_FdADiv_Tube As Integer = 1             ''指定食品A区分 1=流動食
        Public Const TBL_FdADiv_Event As Integer = 2            ''指定食品A区分 2=行事食
        Public Const TBL_FdADiv_TTP As Integer = 3              ''指定食品A区分 3=TTP

        Public Const TBL_StkDiv_Fresh As Integer = 0            ''在庫指定区分 0=生鮮
        Public Const TBL_StkDiv_Zai As Integer = 1              ''在庫指定区分 1=在庫
        Public Const TBL_StkDiv_NotFood As Integer = 2          ''在庫指定区分 2=食品外
        Public Const TBL_StkDiv_NotHac As Integer = 3           ''在庫指定区分 3=発注対象外

        Public Const TBL_HRound_0 As Integer = 0                ''発注数の丸め 0=整数
        Public Const TBL_HRound_1 As Integer = 1                ''発注数の丸め 1=第1位
        Public Const TBL_HRound_2 As Integer = 2                ''発注数の丸め 2=第2位

        Public Const TBL_HCalDiv_RdUp As Integer = 1            ''発注数計算区分 1=切り上げ
        Public Const TBL_HCalDiv_RdDown As Integer = 2          ''発注数計算区分 2=切捨て
        Public Const TBL_HCalDiv_Round As Integer = 3           ''発注数計算区分 3=四捨五入

        Public Const TBL_HCostLv_1 As Integer = 1               ''単価レベル 1=レベル1
        Public Const TBL_HCostLv_2 As Integer = 2               ''単価レベル 2=レベル2

        Public Const TBL_HSalLv_1 As Integer = 1                ''発注販売レベル 1=レベル1
        Public Const TBL_HSalLv_2 As Integer = 2                ''発注販売レベル 2=レベル2

        Public Const TBL_NumOfDiv_None As Integer = 0           ''入り数計算区分 0=なし
        Public Const TBL_NumOfDiv_Avg As Integer = 1            ''入り数計算区分 1=平均
        Public Const TBL_NumOfDiv_Max As Integer = 2            ''入り数計算区分 2=最大
        Public Const TBL_NumOfDiv_Min As Integer = 3            ''入り数計算区分 3=最小

        Public Const TBL_EdSalDiv_None As Integer = 0           ''終売区分 0=終売なし
        Public Const TBL_EdSalDiv_End As Integer = 1            ''終売区分 1=終売

        Public Const TBL_TempZn_Hot As String = "D"             ''温度帯区分 D=常温
        Public Const TBL_TempZn_Cold As String = "C"            ''温度帯区分 C=冷蔵
        Public Const TBL_TempZn_Ice As String = "F"             ''温度帯区分 F=冷凍

        Public Const TBL_ZRptDiv_None As Integer = 0            ''棚卸報告区分 0=報告なし
        Public Const TBL_ZRptDiv_Rept As Integer = 1            ''棚卸報告区分 1=報告

        Public Const TBL_CompDiv_Nor As Integer = 0             ''食品構成区分 0=通常
        Public Const TBL_CompDiv_Comp As Integer = 1            ''食品構成区分 1=構成食品
        Public Const TBL_CompDiv_Parts As Integer = 2           ''食品構成区分 2=パーツ

        Public Const TBL_CompCalDiv_None As Integer = 0         ''食品構成計算区分 0=計算しない
        Public Const TBL_CompCalDiv_Cal As Integer = 1          ''食品構成計算区分 1=計算する

        Public Const TBL_NoSelDiv_Sel As Integer = 0            ''検索除外区分 0=通常
        Public Const TBL_NoSelDiv_NoSel As Integer = 1          ''検索除外区分 1=検索除外

        Public Const TBL_Auth_Jigyo As Integer = 0              ''権限 0=通常
        Public Const TBL_Auth_Siten As Integer = 10             ''権限 10=通常
        Public Const TBL_Auth_Honsya As Integer = 20            ''権限 20=通常
        Public Const TBL_Auth_Kikan As Integer = 30             ''権限 30=通常

        Public Const TBL_SKey_None As Integer = 0               ''保護キー（共通） 0=保護しない
        Public Const TBL_SKey_Protected As Integer = 1          ''保護キー（共通） 1=通常

        Public Const TBL_KonDiv_Kondate As Integer = 0          ''献立区分  0=通常献立
        Public Const TBL_KonDiv_Per As Integer = 1              ''献立区分  1=個人献立

        Public Const TBL_CkTtpDiv_Normal As Integer = 0         ''料理TTP区分  0=通常
        Public Const TBL_CkTtpDiv_Ttp As Integer = 1            ''料理TTP区分  1=TTP

    End Class

    '食事時間
    Public Enum enmEatTimeDiv
        None = 0
        MorningDiv = 10
        AM10Div = 11
        LunchDiv = 20
        PM15Div = 21
        EveningDiv = 30
        NightDiv = 31
        OneDay = 99
    End Enum

    '食品構成区分
    Public Enum enmCompDiv
        Normal = 0    '0:通常食品
        Kousei = 1    '1:食品構成
        Parts = 2     '2:パーツ構成
    End Enum

    'enmSelEtDiv
    Public Enum enmSelEtDiv
        A = 1
        B = 2
        C = 3
        D = 4
        E = 5
        F = 6
        G = 7
        H = 8
        I = 9
        J = 10
        K = 11
        L = 12
        M = 13
        N = 14
        O = 15
        P = 16
        Q = 17
        R = 18
        S = 19
        T = 20
        U = 21
        V = 22
        W = 23
        X = 24
        Y = 25
        Z = 26
    End Enum

    '成分区分
    Public Enum enmNtDiv
        Standard = 1 '基本成分
        Amino = 2    'アミノ酸
        Fatacid = 3  '脂肪酸
        User = 99    'ユーザー
    End Enum

    '端数処理位置
    Public Enum enmRoundPosition
        P1 = 1 '小数点以下1桁を端数処理対象とします。(結果値は小数を持ちません。例:1)
        P2 = 2 '小数点以下2桁を端数処理対象とします。(結果値は小数を最大1桁持ちます。例:1.3)
        P3 = 3 '小数点以下3桁を端数処理対象とします。(結果値は小数を最大2桁持ちます。例:1.31)
        P4 = 4 '小数点以下4桁を端数処理対象とします。(結果値は小数を最大3桁持ちます。例:1.312)
        P5 = 5 '小数点以下5桁を端数処理対象とします。(結果値は小数を最大4桁持ちます。例:1.3123)
        '↓↓Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
        P15 = 15 '小数点以下15桁を端数処理対象とします。(結果値は小数を最大14桁持ちます。)
        '↑↑Add 20110304 FMS)Y.Kurosu Decimal型計算の端数処理対応
    End Enum

    '端数処理区分
    Public Enum enmRoundDiv
        RoundUp = 1 '切り上げ
        CutOff = 2  '切り捨て
        Rounded = 3 '四捨五入
    End Enum

    '端数処理区分2
    Public Enum enmCutUpDiv
        RoundUp = 1 '切り上げ
        CutOff = 2  '切り捨て
    End Enum

    'マスタテーブル
    Public Enum enmMstTables
        SFrm = 1      '画面情報
        SSys = 2      'システム情報
        SMenu = 3     'メニュー情報
        MUser = 4     'ユーザー情報
        MName = 5     '名称マスタ
        MMealTpP = 6  '個人食種マスタ
        MMealTpK = 7  '献立食種マスタ
        MTrader = 8   '仕入先マスタ
        MNamCls = 9   '名称分類マスタ
        MMealGpP = 10 '個人食種グループマスタ
        MMealGpK = 11 '献立食種グループマスタ
        MGpFd = 12    '食品群マスタ(1or2)
        MSisetsu = 13 '施設マスタ
        MGpFdCls = 14 '食品群分類マスタ(1or2)
        MKMealTp = 15 '献立マスタ用食種
        MHeiSt = 16   '併設マスタ

        '＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃
        '＃ どんどん追加して下さい。（･∀･)つ　＃
        '＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃＃

    End Enum

    '名称_成分情報データ区分
    Public Enum enmMNameNtDataDiv
        CdNm = 1   '成分名称
        CdNmAb = 2 '略称
        CdCr = 3   '単位
        CdDic = 4  '小数点桁数
        SortNo = 5 '表示順
    End Enum

    '基幹システム／ＧＴＸ判別フラグ
    Public Enum enmSystem
        K_Sys = 1   '基幹システム
        GTX = 2     'ＧＴＸ
    End Enum

    '有効フラグ
    Public Enum enmValidF
        None = 0    '無効
        Exists = 1  '有効
    End Enum

    '単位計算区分
    Public Enum enmCrCalDiv
        Normal = 0    '通常計算
        Unit = 1      '単位で計算（既存システムのF）
    End Enum

    '**************************************************************************
    '料理マスタ
    '**************************************************************************
    Public Enum enmCkDiv
        Normal = 0  '通常
        MainCok = 1 '主食
        Tube = 2    '経管栄養
        Milk = 3    '調乳   
        Vena = 4    '静脈栄養
    End Enum


    '**************************************************************************
    '食品マスタ
    '**************************************************************************
    '-------------------------------------
    '在庫指定区分
    '-------------------------------------
    Public Enum enmStkDiv
        Fresh = 0       '生鮮
        Stock = 1       '在庫
        NotFood = 2     '食品外
        NotHac = 3      '発注対象外
    End Enum

    '-------------------------------------
    '発注数の丸め
    '-------------------------------------
    Public Enum enmHRound
        AMOUNT_0 = 0    '整数
        AMOUNT_1 = 1    '小数点第１位
        AMOUNT_2 = 2    '小数点第２位
    End Enum

    '-------------------------------------
    '発注数計算区分
    '-------------------------------------
    Public Enum enmHCalDiv
        RoundDUp = 1    '切り上げ
        RoundDown = 2   '切捨て
        Round = 3       '四捨五入
    End Enum

    '-------------------------------------
    '単価レベル
    '-------------------------------------
    Public Enum enmHCostLv
        CostLv1 = 1    'レベル１
        CostLv2 = 2    'レベル２
    End Enum

    '-------------------------------------
    '発注販売レベル
    '-------------------------------------
    Public Enum enmHSalLv
        SalLv1 = 1    'レベル１
        SalLv2 = 2    'レベル２
    End Enum

    '↓ Add by K.Kojima 2008/12/12 Start（１．送信時に調整済のレコードに終売が含まれていた場合のみ警告メッセージを表示する様に変更）*************************************************************************
    '-------------------------------------
    '終売区分	0:終売なし、1:終売食品
    '-------------------------------------
    Public Enum enmEdSalDiv
        None = 0        '終売なし
        SalEnd = 1      '終売食品
    End Enum
    '↑ Add by K.Kojima 2008/12/12 End  （１．送信時に調整済のレコードに終売が含まれていた場合のみ警告メッセージを表示する様に変更）*************************************************************************

    '-------------------------------------
    '棚卸報告区分
    '-------------------------------------
    Public Enum enmZRptDiv
        NotInfo = 0     '報告なし
        Info = 1        '報告あり
    End Enum

    '-------------------------------------
    '権限
    '-------------------------------------
    Public Enum enmAuthority
        Kikan = 30      '基幹システム
        Honsha = 20     '本社
        Shiten = 10     '支店
        Jigyosyo = 0    '事業所
    End Enum


    '-------------------------------------
    '入数区分
    '-------------------------------------
    Public Enum enmNumOfDiv
        NumOfDiv_None = 0 '無し
        NumOfDiv_Avg = 1  '平均
        NumOfDiv_Max = 2  '最大
        NumOfDiv_Min = 3  '最小
    End Enum

    '-------------------------------------
    '送信ステータス
    '-------------------------------------
    Public Enum enmSendSt
        NoData = -1      'データなし
        NoSend = 0       '未送信
        NoPrint = 0      '未印刷
        OneSend = 1      '1回目送信済
        Printded = 1     '印刷済み
        TwoSending = 2   '2回目以降修正中
        TwoSend = 3      '2回目以降送信済
    End Enum

    '**************************************************************************
    '食事時間
    '**************************************************************************
    Public Class EatTime
        Public Const MorningDiv As Integer = 10
        Public Const MorningName As String = "朝食"
        Public Const MorningNameShort As String = "朝"
        Public Const AM10Div As Integer = 11
        Public Const AM10DivName As String = "朝間食"
        Public Const AM10DivNameShort As String = "朝間"
        Public Const LunchDiv As Integer = 20
        Public Const LunchName As String = "昼食"
        Public Const LunchNameShort As String = "昼"
        Public Const PM15Div As Integer = 21
        Public Const PM15Name As String = "昼間食"
        Public Const PM15NameShort As String = "昼間"
        Public Const EveningDiv As Integer = 30
        Public Const EveningName As String = "夕食"
        Public Const EveningNameShort As String = "夕"
        Public Const NightDiv As Integer = 31
        Public Const NightName As String = "夕間食"
        Public Const NightNameShort As String = "夕間"
    End Class

    '-------------------------------------
    '権限（略）
    '-------------------------------------
    Public Const HonAuthNameShort As String = "H"
    Public Const SiAuthNameShort As String = "S"
    Public Const JiAuthNameShort As String = "J"

    '-------------------------------------
    '取扱終了日（期限無し時の設定値 = 移行データの設定値）
    '-------------------------------------

    '７対応 (hayashi)
    Public Const UseEdDtNoLimit As String = "90001231"
    'Public Const UseEdDtNoLimit As String = "20991231"

    '2009/10/28 CM)H.Kuwahara Start 
    'ＣＫＳＫ区分
    Public Enum enmCkSkDiv
        Normal = 0      '通常
        Central = 1     'セントラルキッチン
        Satellite = 2   'サテライトキッチン
    End Enum
    '2009/10/28 CM)H.Kuwahara End 
    Public Const intNutriCnt As Integer = 80

    '-------------------------------------
    '検査区分（栄養指導）
    '-------------------------------------
    Public Class PerExam
        Public Enum enmExamDiv
            ''' <summary>全て</summary>
            ''' <remarks></remarks>
            All = 0

            ''' <summary>栄養指導</summary>
            ''' <remarks></remarks>
            Tec = 1

            ''' <summary>栄養管理</summary>
            ''' <remarks></remarks>
            Mng = 2

            ''' <summary>栄養ケア</summary>
            ''' <remarks></remarks>
            Care = 3
        End Enum

        Public Class MNameCd1
            '-------------------------------------
            '身体計測項目コード（栄養指導）
            '-------------------------------------
            Public Class Body
                Public Const Height As String = "01010"        '身長
                Public Const HeightYmd As String = "01020"     '身長測定日
                Public Const Hizadaka As String = "01030"      '膝高
                Public Const Weight As String = "01040"        '体重
                Public Const WeightYmd As String = "01050"     '体重測定日
                Public Const BMI As String = "01060"           'BMI
            End Class

            '-------------------------------------
            'スクリーニング項目コード（栄養指導）
            '-------------------------------------
            Public Class Scr
                Public Const PrevWeight As String = "02010"    '前回体重
                Public Const PrevWeightYmd As String = "02020" '前回体重測定日
                Public Const WeightFromTo As String = "02030"  '体重測定期間
                Public Const WeightRatio As String = "02040"   '体重変動率
                Public Const Alb As String = "02050"           '血清アルブミン
                Public Const AlbYmd As String = "02060"        '血清アルブミン測定日
                Public Const UseRatio As String = "02070"      '食事摂取量
                Public Const ScrRefCol As String = "02080"     '食事摂取内容
                Public Const Keikou As String = "02090"        '経口栄養
                Public Const Keichou As String = "02100"       '経腸栄養
                Public Const Jyoumyaku As String = "02110"     '静脈栄養
                Public Const Jyokusou As String = "02120"      '褥瘡
                Public Const Other As String = "02130"         'その他
                Public Const Risk As String = "02140"          'リスク
            End Class

            '-------------------------------------
            '栄養補給量算定項目コード（栄養指導）
            '-------------------------------------
            Public Class Nutri
                Public Const Energy As String = "03010"        '必要エネルギー
                Public Const Protein As String = "03020"       '必要タンパク質
                Public Const Fat As String = "03030"           '必要脂肪
                Public Const Carbo As String = "03040"         '必要炭水化物
                Public Const Water As String = "03050"         '必要水分量
                Public Const Salt As String = "03060"          '必要食塩量
                Public Const GiveEnergy As String = "03070"    '給与エネルギー
            End Class

            '-------------------------------------
            '栄養補給量算定項目コード（一般検査項目）
            '-------------------------------------
            Public Class Stand
                Public Const GOT As String = "04010"           'ＧＯＴ
                Public Const GPT As String = "04020"           'ＧＰＴ
                Public Const ALP As String = "04030"           'ＡＬＰ
            End Class

            '-------------------------------------
            '栄養補給量算定項目コード（疾患）
            '-------------------------------------
            Public Class Disease
                Public Const HL As String = "05010"            '高脂血症
                Public Const HLMedic As String = "05020"       '高脂血症・内服
                Public Const DM As String = "05030"            '糖尿病
                Public Const DMMedic As String = "05040"       '糖尿病・内服
                Public Const HT As String = "05050"            '高血圧
                Public Const HTMedic As String = "05060"       '高血圧・内服
                Public Const HD As String = "05070"            '心疾患
                Public Const HDMedic As String = "05080"       '心疾患・内服
            End Class

            '-------------------------------------
            '栄養補給量算定項目コード（その他）
            '-------------------------------------
            Public Class Other
                Public Const Denture As String = "11010"       '義歯
                Public Const PriOpe As String = "11020"        '手術予定
                Public Const OthRefCol As String = "11030"     '備考
            End Class
        End Class

        '-------------------------------------
        '検査項目入力方式（栄養指導）
        '-------------------------------------
        Public Enum enmNtrExamEditor
            ''' <summary>数値入力</summary>
            Dec = 1

            ''' <summary>テキスト入力</summary>
            Txt = 2

            ''' <summary>コンボリスト入力</summary>
            Cmb = 3

            ''' <summary>日付入力</summary>
            Ymd = 4
        End Enum
    End Class

    '**************************************************************************
    '権限
    '**************************************************************************
    '**************************************************************************
    '画面共通定義
    '**************************************************************************
    '**********************************************************************
    '仕入先マスタ
    '**********************************************************************
    Public Class MTraderSec
        Public Const TBL_DEFINE_DLVE As String = "0,当日,-1,前日,-2,前々日" '納品日
    End Class
    '**********************************************************************
    '名称マスタ
    '**********************************************************************
    Public Class MNameDisCd
        Public Const ElementStandard As String = "0001"             '基本成分名
        Public Const ElementAmino As String = "0002"                'アミノ酸成分名
        Public Const ElementFatacid As String = "0003"              '脂肪酸成分名
        Public Const ExDiadetesDiv As String = "0004"               '糖尿交換表区分    
        Public Const ExKidneyDiv As String = "0005"                 '腎臓交換表区分
        Public Const ExDiaKidneyDiv As String = "0006"              '糖尿病腎臓症交換表区分
        Public Const Credit As String = "0010"                      '単位
        Public Const SerchClass As String = "0011"                  '検索分類
        Public Const AppFoodDiv As String = "0012"                  '指定食品区分
        Public Const MethodThawing As String = "0030"               '解凍方法
        Public Const MethodCut As String = "0033"                   'カット方法
        Public Const MethodLowproc As String = "0034"               '下処理方法
        Public Const CookType As String = "0035"                    '料理種類
        Public Const MaterialsType As String = "0036"               '材料種類
        Public Const CookingForm As String = "0037"                 '調理形態
        Public Const Plate As String = "0038"                       '料理皿
        Public Const MealtypeSumPtan1 As String = "0050"            '食種集計パターン１
        Public Const MealtypeSumPtan2 As String = "0051"            '食種集計パターン２
        Public Const MealtypeSumPtan3 As String = "0052"            '食種集計パターン３
        Public Const MealtypeSumPtan4 As String = "0053"            '食種集計パターン４
        Public Const SddishesForm As String = "0054"                '副食形態
        Public Const Sum3Div As String = "0055"                     '集計３区分
        Public Const Sum6Div As String = "0056"                     '集計６区分
        Public Const StsGpCd As String = "0057"                     '統計用食種グループコード
        ' ↓ 2011/06/17 INSERT N.Mimori
        Public Const CommentSumClass As String = "0058"             'コメント形態
        ' ↑ 2011/06/17 INSERT N.Mimori
        Public Const KonForm As String = "0059"                     '献立用形態
        Public Const Tray As String = "0060"                        'トレー
        Public Const MaCkClsCd As String = "0061"                   '主食分類

        '↓Add by K.Ikeda 2011/02/02 (１．名称マスタテーブルの変更に対応)**************************************************
        Public Const MealtypeSumPtan5 As String = "0065"            '食種集計パターン５
        Public Const MealtypeSumPtan6 As String = "0066"            '食種集計パターン６
        Public Const MealtypeSumPtan7 As String = "0067"            '食種集計パターン７
        Public Const MealtypeSumPtan8 As String = "0068"            '食種集計パターン８
        '↑Add by K.Ikeda 2011/02/02 (１．名称マスタテーブルの変更に対応)**************************************************

        Public Const Ward As String = "2001"                        '病棟
        Public Const SickRoom As String = "2002"                    '病室
        Public Const Consulting As String = "2003"                  '診療科
        Public Const LifeCoefficient As String = "2004"             '身体活動レベル
        Public Const UnderReason As String = "2005"                 '欠食理由
        Public Const DeliverPlace As String = "2006"                '配膳先情報
        Public Const SickName As String = "2007"                    '病名
        Public Const BabyBottle As String = "2008"                  '哺乳瓶
        Public Const Nipple As String = "2009"                      '乳首
        Public Const Byoutai As String = "2010"                     '病態食
        Public Const WardClsCd As String = "2011"                   '病棟分類
        Public Const InputPlate As String = "2012"                  '個人摂取率入力皿

        '↓Add by S.Shin 2012/10/18 **************************************************
        Public Const NutMealMethod As String = "2013"               '給与方法
        Public Const NutMealCnt As String = "2014"                  '給与回数
        Public Const NutOneVal As String = "2015"                   '1回量
        '↑Add by S.Shin 2012/10/18 **************************************************

        Public Const MaterialsSumDiv As String = "3001"             '材料集計区分
        Public Const JournalizingDiv As String = "3002"             '仕訳区分
        Public Const Warehouse As String = "4001"                   '倉庫区分

        '↓Add by Hirata 2015/10/07 **************************************************
        Public Const FacDiv As String = "4010"               '工場区分
        '↑Add by Hirata 2015/10/07 **************************************************

        Public Const DistrictName As String = "5001"                '地区名
        Public Const CkClDay As String = "5002"                     'クックチル日付
        Public Const VcCkDay As String = "5003"                     '真空調理日付
        '20111215 S
        Public Const CtlTm As String = "6001"                       '予約時間帯
        Public Const Dietician As String = "6002"                   '栄養士
        Public Const Resv As String = "6003"                        '指導予約枠
        Public Const Tec As String = "6004"                         '指導区分
        Public Const Stress As String = "6005"                      'ストレス係数
        Public Const Active As String = "6006"                      '活動係数
        Public Const ExamCtgry As String = "6007"                   '検査分類
        Public Const ExamItem As String = "6008"                    '検査項目
        Public Const HeightCalcCareMen As String = "6009"           '身長算出係数(栄養ケア式)(男)
        Public Const HeightCalcCareWmn As String = "6010"           '身長算出係数(栄養ケア式)(男)
        Public Const HeightCalcChumMen As String = "6011"           '身長算出係数(Chumlea式)(男)
        Public Const HeightCalcChumWmn As String = "6012"           '身長算出係数(Chumlea式)(女)
        Public Const BeeHarrisMen As String = "6013"                'BEE係数(ハリス・ベネディクト式)(男)
        Public Const BeeHarrisWmn As String = "6014"                'BEE係数(ハリス・ベネディクト式)(女)
        Public Const BeeJpnMen As String = "6015"                   'BEE係数(日本人簡易式)(男)
        Public Const BeeJpnWmn As String = "6016"                   'BEE係数(日本人簡易式)(女)
        Public Const Risk As String = "6017"                        'リスク（コンボボックス用）
        Public Const PlusMinus As String = "6018"                   '陽性陰性（コンボボックス用）
        Public Const Exists As String = "6019"                      '有無（コンボボックス用）
        Public Const WeightFromTo As String = "6020"                '体重計測期間
        '20111215 E
        Public Const Nameofera As String = "9001"                   '年号
        Public Const SalesTax As String = "9002"                    '消費税

        ' ↓ 2011/06/17 DELETE N.Mimori
        ''7/7現在、UclSelectFrm.SetItemNameCommentが未対応の為、対応完了後に削除予定
        ''↓ Delete by K.Kojima 2008/07/07 Start（１.コメント集計分類は不要な為、削除）*************************************************************************
        'Public Const CommentSumClass As String = "0058"             'コメント集計分類
        ''↑ Delete by K.Kojima 2008/07/07 End  （１.コメント集計分類は不要な為、削除）*************************************************************************
        ' ↑ 2011/06/17 DELETE N.Mimori
    End Class

    '**********************************************************************
    '個人情報
    '**********************************************************************
    Public Class PerCST

        '個人情報用コンスト
        Public Const MINVALUE_T_PerYMD As String = "00000000"           '個人情報の日付最大値（値なし意味）
        Public Const MINVALUE_T_PerTmDiv As Integer = 0                 '個人情報の時間区分最大値（値なしの意味）

        Public Const LIMITVALUE_T_PerYMD As String = "99999999"         '個人情報の日付最大値（～ずっとの意味）
        Public Const LIMITVALUE_T_PerTmDiv As Integer = 99              '個人情報の時間区分最大値（～ずっとの意味）

        '=======================================================
        '2009/01/14 FTC)T.IKEDA 背景色をBrown⇒DarkGrayに変更
        '=======================================================
        Public Const NO_MEAL_BACKCOLOR_NAME As String = "DarkGray"
        Public Const NO_MEAL_FORECOLOR_NAME As String = "White"

        Public Const PERCD_LEN As Integer = 10                          '個人コード桁数

        '欠食区分
        Public Enum enmNoMeal
            Normal = 0  '通常
            NoMeal = 1  '欠食
        End Enum

        '個人献立区分
        Public Enum enmPerMn
            Normal = 0  '通常献立
            PerMn = 1   '個人献立
        End Enum
        '個人献立保護区分
        Public Enum enmPerProtect
            Normal = 0      'なし
            Protect = 1     'あり
        End Enum

        '献立数区分
        Public Enum enmKCntDiv
            Normal = 0  '通常
            Week = 1    '週間
        End Enum

        '性別
        Public Enum enmSex
            Man = 1     '男
            Woman = 2   '女
        End Enum

        '削除除外区分
        Public Enum enmUnDelDiv
            Normal = 0  '通常
            NoDel = 1   '削除除外
        End Enum

        '特別食加算区分
        Public Enum enmEatAdDiv
            NoAdd = 0   '非加算
            Add = 1     '加算
        End Enum

        '食札区分
        Public Enum enmMCPrt
            None = 0    'なし
            Exists = 1  '有り
        End Enum

        '指導区分
        Public Enum enmTecDiv
            Personal = 1   '個人
            Group = 2      '集団
            Outpatient = 3 '外来
        End Enum

        'NST対応
        Public Enum enmNstDiv
            NoNst = 0     'NSTなし
            Nst = 1   'NST対応
        End Enum

        ''延食
        Public Const DELIVERPLACE_EXTENSION_MEAL_CODE As String = "77777"

    End Class

    '↓ Add by K.Kojima 2008/12/05 Start（１．仕入区分を表示する様に変更）*************************************************************************
    '**********************************************************************
    '仕入情報情報
    '**********************************************************************
    Public Class BuyCST

        '仕入区分
        Public Enum enmStkDiv
            Fresh = 0       '生鮮食品（在庫管理不要）
            Zai = 1         '在庫食品
            NotFood = 2     '食品外（バラン等）
            NotHac = 3      '発注対象外
        End Enum

    End Class
    '↑ Add by K.Kojima 2008/12/05 End  （１．仕入区分を表示する様に変更）*************************************************************************

    '**********************************************************************
    '棚卸マスタ
    '**********************************************************************
    Public Enum enmZaiDiv
        Patient = 1     '個人
        Other = 2       '個人外
    End Enum

    '**********************************************************************
    '発注業務
    '**********************************************************************
    Public Class THacInf

        '↓ Change by K.Kojima 2008/09/18 Start（１．過去の納品日はエラー、当日納品日は警告、未来の納品日は正常となる様に仮対応）*************************************************************************
        Public Const DATA_NO_SEND_DAY As Integer = 0                'データ伝送不可日数（当日からデータ伝送ができない日数を設定）
        'Public Const DATA_NO_SEND_DAY As Integer = 1                'データ伝送不可日数（当日からデータ伝送ができない日数を設定）
        '↑ Change by K.Kojima 2008/09/18 End  （１．過去の納品日はエラー、当日納品日は警告、未来の納品日は正常となる様に仮対応）*************************************************************************

        '配信済フラグ
        Public Enum enmSndDiv
            NoDelivery = 0  '未配信
            Delivery = 1    '配信済
        End Enum

        Public Const CST_HAC_DEADLINE_FILE As String = "Aji7HDF"
        Public Const CST_HAC_MEQSSY_FILE As String = "Aji7MekyuHac.csv"

        ''' <summary>
        ''' 工場区分-現地判定文字列(DisCd=4010,Cd1がZZで始まるものは現地とする)
        ''' </summary>
        ''' <remarks>工場区分-現地判定文字列</remarks>
        Public Const CST_FACDIV_GENTHI_START_CHAR As String = "ZZ"
    End Class

    '併設事業所区分
    Public Enum enmHeiStDiv
        Normal = 0      '通常
        Heisetsu = 1    '併設
        Kinrin = 2      '近隣
    End Enum

    '**********************************************************************
    ' 集配信データ区分
    '**********************************************************************
    Public Class SendRecvDiv
        Public Const SRDiv_Memo_G As String = "GMEM"    'メモ
    End Class

    '↓ Add by K.Kojima 2009/01/06 Start（１．マスタ送受信時に内容を表示する機能を追加）*************************************************************************
    '**************************************************************************
    '最多食品群マスタ
    '**************************************************************************
    Public Enum enmGpFdDiv
        First = 1    '第1食品群
        Second = 2   '第2食品群
    End Enum

    '**************************************************************************
    '送受信マスタ詳細画面用
    '**************************************************************************
    Public Class DetailCST

        '**************************************************************************
        '画面判別文字
        '**************************************************************************
        Public Const CST_DISTINCTION As String = "FrmMstDetail"

        '**************************************************************************
        '区切り文字
        '**************************************************************************
        Public Const CST_DELIMITER As String = "_"

        '**************************************************************************
        '送受信判別文字
        '**************************************************************************
        Public Class StrSndRcvCST
            Public Const Send As String = "S"           '送信
            Public Const Receive As String = "R"        '受信
        End Class

        '**************************************************************************
        'マスタ判別文字
        '**************************************************************************
        Public Class StrMstCST
            Public Const MFd As String = "MFd"          '食品マスタ
            Public Const MNutri As String = "MNutri"    '成分マスタ
            Public Const MCok As String = "MCok"        '料理マスタ
        End Class

    End Class
    '↑ Add by K.Kojima 2009/01/06 End  （１．マスタ送受信時に内容を表示する機能を追加）*************************************************************************

    '↓hayashi 2013/04/03
    '**************************************************************************
    'ベンダーコード
    '**************************************************************************
    Public Class VendorCd
        Public Const NGF As String = "005" '日本ゼネラルフード(NGF)
    End Class

    '**************************************************************************
    '仕入先マスタの指定食品区分(NGFにて使用)
    '**************************************************************************
    Public Class FdDiv
        Public Const R0 As String = "00000" 'なし
        Public Const R1 As String = "00001" 'R1
        Public Const R2 As String = "00002" 'R2
        Public Const R3 As String = "00003" 'R3
    End Class
    '↑hayashi 2013/04/03

    '**********************************************************************
    ' GtUpdater機能
    '**********************************************************************
    Public Class GtUpdater
        '**********************************************************************
        ' SQLCMD用パラメータ
        '**********************************************************************
        Public Const C_SQLSERVER_SCMD As String = "SQLCMD.exe"
        Public Const C_SQLSERVER_USER As String = "-U{0}"
        Public Const C_SQLSERVER_PASSWORD As String = "-P{0}"
        Public Const C_SQLSERVER_DBNAME As String = "-d{0}"
        Public Const C_SQLSERVER_DBSERVER As String = "-S{0}"
        Public Const C_SQLSERVER_COMMAND As String = "-i{0}"

        '最新版｢GtUpdater｣格納場所
        Public Const GtUpdaterFD As String = "updater"

        ' データカタログ情報列挙型
        Public Enum enumDtClg
            DBServer = 0        ' DBServerName
            DBName = 1          ' DBName
            UserID = 2          ' UserID
            Password = 3        ' Password
            Max = 3             ' 最大インデクス番号
        End Enum

    End Class

    '**********************************************************************
    ' バックアップ機能
    '**********************************************************************
    Public Class Backup

        '**********************************************************************
        'ＸＭＬセクション名
        '**********************************************************************
        Public Class xmlSectionName
            Public Const PathSection As String = "PathSection"         'パスセクション
            Public Const application As String = "application"         'アプリケーションセクション
        End Class

        '**********************************************************************
        'ＸＭＬセクションタグ名
        '**********************************************************************
        Public Class xmlTagName
            Public Const InitialCatalog As String = "InitialCatalog"    'イニシャルカタログ
            Public Const DataSource As String = "DataSource"            'データソース
            Public Const User As String = "User"                        'ユーザ
            Public Const Password As String = "Password"                'パスワード
            Public Const ConnectTimeout As String = "ConnectTimeout"    '接続タイムアウト
            Public Const CommandTimeout As String = "CommandTimeout"    'コマンドタイムアウト
            Public Const Pooling As String = "Pooling"                  '接続プール設定
            Public Const BackupSakiPath As String = "BackupSakiPath"    '保存先フォルダ初期値
            Public Const RestoreMotoPath As String = "RestoreMotoPath"  '復元元フォルダ初期値
            Public Const ZipNameSoft As String = "ZipNameSoft"          'ソフト圧縮ファイル名初期値
            Public Const ZipNameDB As String = "ZipNameDB"              'DB圧縮ファイル名初期値
            'Public Const TempPath As String = "TempPath"                'テンポラリフォルダ作成パス名
            Public Const Expansion As String = "Expansion"              '圧縮ファイル拡張子名
            Public Const SoftBackName As String = "SoftBackName"        'ソフトバックアップ名
            Public Const DebugFlg As String = "DebugFlg"                'デバック環境フラグ
            Public Const SoftPath As String = "SoftPath"                'ソフト存在フォルダパス
            Public Const ServiceName As String = "ServiceName"          'サービス名
            'Public Const DBSavePlace As String = "DBSavePlace"          'データベースファイル保存場所
            Public Const BakExeFileName As String = "BakExeFileName"    'バックアップEXEファイル名
            Public Const BackupFolder As String = "BackupFolder"        'バックアップファイル保存フォルダ
            Public Const CmdLineStr As String = "CmdLineStr"            '起動元判別コマンドライン文字列
            Public Const ResExeFileName As String = "ResExeFileName"    'リストアEXEファイル名
        End Class

    End Class

    '**********************************************************************
    'テーブル名
    '**********************************************************************
    Public Class TableName
        Public Const M_Fd As String = "M_Fd"                    '食品マスタ
        Public Const M_FdSKey As String = "M_FdSKey"            '食品保護キーマスタ
        Public Const M_FdCost As String = "M_FdCost"            '食品単価マスタ
        Public Const M_FdCmt As String = "M_FdCmt"              '食品コメントマスタ
        Public Const M_FdComp As String = "M_FdComp"            '食品構成マスタ
        Public Const M_FdNoUse As String = "M_FdNoUse"          '使用見合わせ食品マスタ
        Public Const M_Nutri As String = "M_Nutri"              '成分マスタ
        Public Const M_NutSKey As String = "M_NutSKey"          '成分保護キーマスタ
        Public Const M_Trader As String = "M_Trader"            '仕入先マスタ
        Public Const M_Cmt As String = "M_Cmt"                  'コメントマスタ
        Public Const M_CmtChg As String = "M_CmtChg"            'コメント変換マスタ
        Public Const M_NamCls As String = "M_NamCls"            '名称分類マスタ
        Public Const M_Name As String = "M_Name"                '名称マスタ
        Public Const M_ContCls As String = "M_ContCls"          '契約分類マスタ
        Public Const M_ContClsLk As String = "M_ContClsLk"      '契約分類関連マスタ
        Public Const M_Cok As String = "M_Cok"                  '料理マスタ
        Public Const M_CokFd As String = "M_CokFd"              '料理食品マスタ
        Public Const M_MealTpP As String = "M_MealTpP"          '食種マスタ（個人）
        Public Const M_MealTpK As String = "M_MealTpK"          '食種マスタ（献立）
        Public Const M_MealZai As String = "M_MealZai"          '食種材区マスタ
        Public Const M_MealLmtP As String = "M_MealLmtP"        '食種制限値マスタ(個人）
        Public Const M_MealLmtK As String = "M_MealLmtK"        '食種制限値マスタ(献立）
        Public Const M_MealCmtP As String = "M_MealCmtP"        '食種コメントマスタ（個人）
        Public Const M_MealGpP As String = "M_MealGpP"          '食種グループマスタ（個人）
        Public Const M_MealGpK As String = "M_MealGpK"          '食種グループマスタ（献立）
        Public Const M_GpFdClsMx As String = "M_GpFdClsMx"      '最多食品群分類マスタ
        Public Const M_GpFdMx As String = "M_GpFdMx"            '最多食品群マスタ
        Public Const M_GpFdCls1 As String = "M_GpFdCls1"        '食品群分類マスタ1
        Public Const M_GpFd1 As String = "M_GpFd1"              '食品群マスタ1
        Public Const M_GpFdLk1 As String = "M_GpFdLk1"          '食品群関連マスタ1
        Public Const M_GpFdCls2 As String = "M_GpFdCls2"        '食品群分類マスタ2
        Public Const M_GpFd2 As String = "M_GpFd2"              '食品群マスタ2
        Public Const M_GpFdLk2 As String = "M_GpFdLk2"          '食品群関連マスタ2
        Public Const M_Age As String = "M_Age"                  '年齢管理マスタ
        Public Const M_AgeInf As String = "M_AgeInf"            '年齢情報マスタ
        Public Const M_AgeDNR As String = "M_AgeDNR"            '年齢所要量マスタ
        Public Const T_FdGpAvgK As String = "T_FdGpAvgK"        '食糧構成（加重）
        Public Const T_FdGpAvgJ As String = "T_FdGpAvgJ"        '食糧構成（実施）
        Public Const M_SiTen As String = "M_SiTen"              '支店マスタ
        Public Const M_JiGyo As String = "M_JiGyo"              '事業所マスタ
        Public Const M_Sisetsu As String = "M_Sisetsu"          '施設マスタ
        Public Const M_HeiSt As String = "M_HeiSt"              '併設（近隣）情報マスタ
        Public Const M_SpDays As String = "M_SpDays"            '記念日マスタ
        Public Const M_KMealTp As String = "M_KMealTp"          '献立マスタ用食種
        Public Const M_KMealLmt As String = "M_KMealLmt"        '献立マスタ用食種制限値
        Public Const M_KMealLk As String = "M_KMealLk"          '献立マスタ用食種関連マスタ
        Public Const M_User As String = "M_User"                'ユーザーマスタ
        Public Const M_UserGp As String = "M_UserGp"            'ユーザーグループマスタ
        Public Const M_Color As String = "M_Color"              '色情報マスタ
        Public Const M_ColorSet As String = "M_ColorSet"        '色情報設定マスタ
        Public Const T_MKonInf As String = "T_MKonInf"          '献立マスタ
        Public Const T_MKonMk As String = "T_MKonMk"            '献立マスタ作成
        Public Const T_MKonCk As String = "T_MKonCk"            '献立料理マスタ
        Public Const T_MKonFd As String = "T_MKonFd"            '献立食品マスタ
        Public Const T_KonMk As String = "T_KonMk"              '通常献立作成状況
        Public Const T_KonCk As String = "T_KonCk"              '通常献立料理
        Public Const T_KonFd As String = "T_KonFd"              '通常献立食品
        Public Const T_CKonCk As String = "T_CKonCk"            '代替献立料理
        Public Const T_CKonFd As String = "T_CKonFd"            '代替献立食品
        Public Const T_CKonCnt As String = "T_CKonCnt"          '代替献立人数
        Public Const T_PerM As String = "T_PerM"                '個人マスタ
        Public Const T_PerHospi As String = "T_PerHospi"        '個人入退院情報
        Public Const T_PerInf As String = "T_PerInf"            '個人情報
        Public Const T_PerEat As String = "T_PerEat"            '個人食事情報
        Public Const T_PerKCnt As String = "T_PerKCnt"          '個人献立数情報
        Public Const T_PerMov As String = "T_PerMov"            '個人移動情報
        Public Const T_PerDlv As String = "T_PerDlv"            '個人配膳情報
        Public Const T_PerCmt As String = "T_PerCmt"            '個人コメント
        Public Const T_PerLmt As String = "T_PerLmt"            '個人制限値
        Public Const T_PerMn As String = "T_PerMn"              '個人献立情報
        Public Const T_PerSel As String = "T_PerSel"            '個人選択食
        Public Const T_PerNut As String = "T_PerNut"            '個人経管栄養
        Public Const T_PerMik As String = "T_PerMik"            '個人調乳
        Public Const T_PerVena As String = "T_PerVena"          '個人静脈栄養
        Public Const T_PerIntk As String = "T_PerIntk"          '個人摂取値
        Public Const T_PerTec As String = "T_PerTec"            '個人栄養指導
        Public Const T_PerIll As String = "T_PerIll"            '個人病名
        Public Const T_MPCnt As String = "T_MPCnt"              '予定実施人数
        Public Const T_HacInf As String = "T_HacInf"            '発注情報
        Public Const T_HacJZ As String = "T_HacJZ"              '発注準在庫情報
        Public Const T_HacFd As String = "T_HacFd"              '発注食品情報
        Public Const T_HacHstNo As String = "T_HacHstNo"        '発注送信履歴枝番
        Public Const T_HacHst As String = "T_HacHst"            '発注送信履歴
        Public Const T_Buy As String = "T_Buy"                  '仕入情報
        Public Const T_BuyInf As String = "T_BuyInf"            '仕入食品情報
        Public Const T_ZaiOut As String = "T_ZaiOut"            '在庫払出情報
        Public Const T_Invty As String = "T_Invty"              '棚卸情報
        Public Const T_ZaiMony As String = "T_ZaiMony"          '在庫簡易金額情報
        Public Const T_ZaiSim As String = "T_ZaiSim"            '在庫簡易情報
        Public Const R_HVirusC As String = "R_HVirusC"          '加熱・殺菌加工記録簿料理
        Public Const R_HVirusF As String = "R_HVirusF"          '加熱・殺菌加工記録簿食品
        Public Const R_HSeq As String = "R_HSeq"                '発注書連番管理
        Public Const R_EiyoS As String = "R_EiyoS"              '栄養出納年報
        Public Const W_CvtCd As String = "W_CvtCd"              '移行コード変換WK
        Public Const W_5Nutri As String = "W_5Nutri"            '5訂成分ワーク
        Public Const S_Frm As String = "S_Frm"                  '画面マスタ
        Public Const S_RptInf As String = "S_RptInf"            '帳票マスタ
        Public Const S_ROutPut As String = "S_ROutPut"          '帳票印刷情報マスタ
        Public Const S_Menu As String = "S_Menu"                'メニューマスタ
        Public Const S_Msg As String = "S_Msg"                  'メッセージマスタ
        Public Const S_Sys As String = "S_Sys"                  'システム設定マスタ
        Public Const S_SCK As String = "S_SCK"                  'ショートカットキーマスタ
        Public Const S_Ctl As String = "S_Ctl"                  'コントロール設定マスタ
        Public Const S_JigyoSnd As String = "S_JigyoSnd"        '事業所設定画面マスタ
    End Class

#Region "自動付番定義"
    '**************************************************************************
    ' 自動付番定義
    ' ※付番初期値、付番制限値はフィールドの桁数で0埋めした1以上の数値を記述して下さい。
    ' ★付きのテーブルは、ログイン時の施設コードを引き渡す必要があります。    
    '**************************************************************************
    Public Class AutoNumSec
        'テーブル名称
        Public Const DBNAME_S_Menu As String = "S_Menu"             'メニューマスタ
        Public Const DBNAME_M_Sisetsu As String = "M_Sisetsu"       '施設マスタテーブル名(M_Sisetsu)
        Public Const DBNAME_M_Trader As String = "M_Trader"         '仕入先マスタテーブル名(M_Trader)
        '----------------------------------------------------------------------------------------------自動振番対応
        'Public Const DBNAME_M_Trader_H As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'H%')"         '仕入先マスタテーブル名(本社用)	2008/07/14
        'Public Const DBNAME_M_Trader_S As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'S%')"         '仕入先マスタテーブル名(支店用)	2008/07/14
        'Public Const DBNAME_M_Trader_J As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'J%')"         '仕入先マスタテーブル名(事業所用)	2008/07/14


        Public Const DBNAME_M_Fd As String = "M_Fd"                 '食品マスタテーブル
        'Public Const DBNAME_M_Fd_H As String = "(select substring(FdCd,2,7) as FdCd from M_Fd where substring(FdCd,1,1) = 'H')"       '食品マスタテーブル(本社用)   (通常)     2009/09/11 追加
        'Public Const DBNAME_M_Fd_S As String = "(select substring(FdCd,2,7) as FdCd from M_Fd where substring(FdCd,1,1) = 'S')"       '食品マスタテーブル(支店用)   (通常)     2009/09/11 追加
        Public Const DBNAME_M_MealTpK As String = "M_MealTpK"       '食種マスタ(献立)テーブル
        Public Const DBNAME_M_MealTpP As String = "M_MealTpP"       '食種マスタ(個人)テーブル
        Public Const DBNAME_M_MealGpK As String = "M_MealGpK"       '食種グループマスタ(献立)テーブル
        Public Const DBNAME_M_MealGpP As String = "M_MealGpP"       '食種グループマスタ(食数)テーブル
        Public Const DBNAME_M_KMealTp As String = "M_KMealTp"       '献立マスタ用食種テーブル
        'Public Const DBNAME_M_KMealTp_H As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'H')"       '献立マスタ用食種テーブル(本社用)
        'Public Const DBNAME_M_KMealTp_S As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'S')"       '献立マスタ用食種テーブル(支店用)
        'Public Const DBNAME_M_KMealTp_J As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'J')"       '献立マスタ用食種テーブル(事業所用)
        Public Const DBNAME_M_Cok As String = "M_Cok"               '料理マスタテーブル
        'Public Const DBNAME_M_Cok_H As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'H')"       '料理マスタテーブル(本社用)   (通常)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_S As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'S')"       '料理マスタテーブル(支店用)   (通常)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_J As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'J')"       '料理マスタテーブル(事業所用) (通常)     2008/06/16 追加

        'Public Const DBNAME_M_Cok_HSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HSYU')" '料理マスタテーブル(本社用)   (主食)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_SSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SSYU')" '料理マスタテーブル(支店用)   (主食)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_JSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JSYU')" '料理マスタテーブル(事業所用) (主食)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_HKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HKKN')" '料理マスタテーブル(本社用)   (経管栄養) 2008/06/16 追加
        'Public Const DBNAME_M_Cok_SKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SKKN')" '料理マスタテーブル(支店用)   (経管栄養) 2008/06/16 追加
        'Public Const DBNAME_M_Cok_JKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JKKN')" '料理マスタテーブル(事業所用) (経管栄養) 2008/06/16 追加
        'Public Const DBNAME_M_Cok_HMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HMLK')" '料理マスタテーブル(本社用)   (調乳)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_SMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SMLK')" '料理マスタテーブル(支店用)   (調乳)     2008/06/16 追加
        'Public Const DBNAME_M_Cok_JMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JMLK')" '料理マスタテーブル(事業所用) (調乳)     2008/06/16 追加

        Public Const DBNAME_M_Cok_SYU As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'SYU')" '料理マスタテーブル(主食)     2011/12/15 追加
        Public Const DBNAME_M_Cok_KKN As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'KKN')" '料理マスタテーブル(経管栄養) 2011/12/15 追加
        Public Const DBNAME_M_Cok_MLK As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'MLK')" '料理マスタテーブル(調乳)     2011/12/15 追加
        Public Const DBNAME_M_Cok_VEN As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'VEN')" '料理マスタテーブル(静脈栄養)     2011/12/15 追加


        Public Const DBNAME_M_Age As String = "M_Age"               '年齢管理マスタテーブル
        'Public Const DBNAME_M_Age_H As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'H')"       '年齢管理マスタテーブル(本社用)
        'Public Const DBNAME_M_Age_S As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'S')"       '年齢管理マスタテーブル(支店用)
        'Public Const DBNAME_M_Age_J As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'J')"       '年齢管理マスタテーブル(事業所用)
        Public Const DBNAME_T_PerM As String = "T_PerM"             '個人マスタ
        Public Const DBNAME_M_Nutri As String = "M_Nutri"           '成分マスタ
        'Public Const DBNAME_M_Nutri_H As String = "(select substring(NtCd,2,7) as NtCd,NtDiv from M_Nutri where substring(NtCd,1,1) = 'H')"   '成分マスタテーブル(本社用)   (通常)     2009/09/11 追加
        'Public Const DBNAME_M_Nutri_S As String = "(select substring(NtCd,2,7) as NtCd,NtDiv from M_Nutri where substring(NtCd,1,1) = 'S')"   '成分マスタテーブル(支店用)   (通常)     2009/09/11 追加
        Public Const DBNAME_A_Config As String = "A_Config"

        'フィールド名
        Public Const FIELDNAME_MENUNO As String = "MenuNo"              'フィールド名
        Public Const FIELDNAME_JICD As String = "JiCd"                  'フィールド名（JiCd）
        Public Const FIELDNAME_TDCD As String = "TrndCd"                'フィールド名（TdCd）
        Public Const FIELDNAME_FD_FDCD As String = "FdCd"               'フィールド名
        Public Const FIELDNAME_MEALTPK_MEALCD As String = "MealCd"      'フィールド名
        Public Const FIELDNAME_MEALTPP_MEALCD As String = "MealCd"      'フィールド名
        Public Const FIELDNAME_MMealGpK_MealGpCd As String = "MealGpCd" 'フィールド名：食種グループコード
        Public Const FIELDNAME_MMEALGPP_MEALGPCD As String = "MealGpCd" 'フィールド名：食種グループコード
        Public Const FIELDNAME_CKCD As String = "CkCd"                  'フィールド名（CkCd）
        Public Const FIELDNAME_MAGE_AGECD As String = "AgeCd"           'フィールド名（AgeCd）
        Public Const FIELDNAME_TPerM_PerCD As String = "PerCd"          'フィールド名(個人コード)
        Public Const FIELDNAME_MKMEALTP_MEALCD As String = "MealCd"     'フィールド名
        Public Const FIELDNAME_NtCd As String = "NtCd"                  'フィールド名（成分コード）

        '付番初期値
        Public Const INITVALUE_S_MENU As String = "0001"              'メニューマスタ付番初期値
        Public Const INITVALUE_M_SISETSU As String = "0001"           '施設マスタ付番初期値

        Public Const INITVALUE_M_TRADER As String = "00000001"         '仕入先マスタ付番初期値Aji7対応
        'Public Const INITVALUE_M_TRADER As String = "0000001"         '仕入先マスタ付番初期値(変更の可能性あり)
        Public Const LIMITVALUE_M_TRADER As String = "99999999"        '仕入先マスタ付番制限値Aji７対応８桁に変更
        'Public Const LIMITVALUE_M_TRADER As String = "9999999"        '仕入先マスタ付番制限値(変更の可能性あり)頭に[HSJ]を付与するために7桁に変更

        Public Const INITVALUE_M_FD As String = "00000001"            '食品マスタ付番初期値(変更の可能性あり)
        Public Const LIMITVALUE_M_FD As String = "99999999"           '食品マスタ付番制限値(変更の可能性あり)

        Public Const INITVALUE_M_MEALTPK As String = "0001"           '食種マスタ(献立)付番初期値(変更の可能性あり)
        Public Const INITVALUE_M_MEALTPP As String = "0001"           '食種マスタ(個人)付番初期値(変更の可能性あり)
        Public Const INITVALUE_M_MealGpK As String = "0001"           '食種グループマスタ(献立)付番初期値(変更の可能性あり)
        Public Const INITVALUE_M_MEALGPP As String = "0001"           '食種グループマスタ(食数)付番初期値(変更の可能性あり)

        Public Const INITVALUE_M_Cok As String = "000000001"           '料理マスタ付番初期値
        'Public Const INITVALUE_M_Cok As String = "00000001"           '料理マスタ付番初期値

        Public Const INITVALUE_M_CokNotNormal As String = "000001"     '料理マスタ付番初期値(通常以外) 2011/12/15 追加
        'Public Const INITVALUE_M_CokNotNormal As String = "00001"     '料理マスタ付番初期値(通常以外) 2008/06/16 追加

        Public Const INITVALUE_M_AGE As String = "00001"              '年齢管理マスタ付番初期値
        Public Const INITVALUE_M_AGE_AUTH As String = "0001"          '年齢管理マスタ付番初期値(権限) 2008/06/17 追加
        Public Const INITVALUE_T_PerM As String = "0000000001"        '個人マスタ付番初期値(変更の可能性あり)
        Public Const INITVALUE_M_KMEALTP As String = "0001"           '献立マスタ用食種
        Public Const INITVALUE_M_KMEALTP_AUTH As String = "001"       '献立マスタ用食種(権限)
        Public Const INITVALUE_M_NtCd As String = "00000001"          '成分マスタ付番初期値

        '付番制限値
        Public Const LIMITVALUE_M_S_MENU As String = "9999"           'メニューマスタ付番制限値
        Public Const LIMITVALUE_M_SISETSU As String = "9999"          '施設マスタ付番制限値



        Public Const LIMITVALUE_M_MEALTPK As String = "9999"          '食種マスタ(献立)付番制限値(変更の可能性あり)
        Public Const LIMITVALUE_M_MEALTPP As String = "9999"          '食種マスタ(個人)付番制限値(変更の可能性あり)
        Public Const LIMITVALUE_M_MealGpK As String = "9999"          '食種グループマスタ(献立)付番制限値(変更の可能性あり)
        Public Const LIMITVALUE_M_MEALGPP As String = "9999"          '食種グループマスタ(食数)付番制限値(変更の可能性あり)


        Public Const LIMITVALUE_M_Cok As String = "999999999"          '料理マスタ付番制限値
        Public Const LIMITVALUE_M_CokNotNormal As String = "999999"    '料理マスタ付番制限値(通常以外) 2008/06/16 追加

        'Public Const LIMITVALUE_M_Cok As String = "99999999"          '料理マスタ付番制限値
        'Public Const LIMITVALUE_M_CokNotNormal As String = "99999"    '料理マスタ付番制限値(通常以外) 2008/06/16 追加

        Public Const LIMITVALUE_M_AGE As String = "99999"             '年齢管理マスタ付番制限値
        Public Const LIMITVALUE_M_AGE_AUTH As String = "9999"         '年齢管理マスタ付番制限値(権限) 2008/06/17 追加
        Public Const LIMITVALUE_T_PerM As String = "9999999999"       '個人マスタ付番制限値(変更の可能性あり)

        Public Const LIMITVALUE_M_KMEALTP As String = "9999"          '献立マスタ用食種
        Public Const LIMITVALUE_M_KMEALTP_AUTH As String = "999"      '献立マスタ用食種(権限)
        Public Const LIMITVALUE_M_NtCd As String = "99999999"         '成分マスタ付番制限値

    End Class

#End Region

#Region "帳票印字関連"
    '**********************************************************************
    '帳票印字関連
    '**********************************************************************
    Public Class Report

        Public Const PRINT_EXTENSION_TXT As String = ".txt"         'txtファイル
        Public Const PRINT_EXTENSION_CSV As String = ".csv"         'csvファイル
        Public Const PRINT_EXTENSION_RPT As String = ".rpt"         'rptファイル
        Public Const PRINT_EXTENSION_PDF As String = ".pdf"         'pdfファイル
        Public Const PRINT_EXTENSION_XLS As String = ".xlsm"        'xlsファイル
        Public Const PRINT_EXTENSION_ZIP As String = ".zip"         'zipファイル
        Public Const PRINT_EXTENSION_EXE As String = ".exe"         'exeファイル
        Public Const PRINT_EXTENSION_BAT As String = ".bat"         'batファイル
        'Public Const PRINT_EXTENSION_A7R As String = ".a7r"         'a7rファイル
        Public Const PRINT_EXTENSION_A7U As String = ".a7u"         'a7uファイル(アップロード中)
        Public Const PRINT_EXTENSION_A7D As String = ".a7d"         'a7dファイル(ダウンロード中)

        Public Const REPORT_FILE_HEADER_NAME As String = "Header"   'CSVのヘッダーファイル名
        Public Const REPORT_FILE_DELIMITER As String = ">>>>> This line is Delimiter of ReportData for AjiDAS <<<<<"   '重複レポートファイルのデリミタ
        Public Const REPORT_FILE_ERROR_INF As String = ">>>>> Error Information <<<<<"

        Public Const REPORT_FILE_MDB_NAME As String = "GTXPrint.mdb"
        Public Const REPORT_FILE_SCHEMA_NAME As String = "schema.ini"

    End Class
#End Region

#Region "オーダ連携関連"

    Public Class ORecv

        Public Const ORDER_DEFINE_ORDERTYPE_KAMEDA As String = "kameda"         '亀田オーダ
        Public Const ORDER_DEFINE_ORDERTYPE_MIRAIS As String = "MIRAIs"         'MIRAIs（NEC・CSI）
        Public Const ORDER_DEFINE_ORDERTYPE_FUJITSU_GX As String = "GX"         'GX（富士通）
        Public Const ORDER_DEFINE_ORDERTYPE_SOFTS As String = "SoftS"           '会津(SoftS)
        Public Const ORDER_DEFINE_ORDERTYPE_PHRMAKER As String = "PHRMAKER"     'PHRMAKER(NJC)
        Public Const ORDER_DEFINE_ORDERTYPE_SBS As String = "SBS"               'SBS
        Public Const ORDER_DEFINE_ORDERTYPE_SOFT_STANDARD As String = "SoftD"   'SoftSの標準
        Public Const ORDER_DEFINE_ORDERTYPE_NAIS As String = "NAIS"             'ナイス
        Public Const ORDER_DEFINE_ORDERTYPE_ALICE As String = "ALICE"           'アリス
        Public Const ORDER_DEFINE_ORDERTYPE_FUJITSU_EX_WEBEDITION As String = "EXWEB"   'EX WebEdition（富士通）
        Public Const ORDER_DEFINE_ORDERTYPE_MIC As String = "MIC"               'MIC メディカル情報（株）
        Public Const ORDER_DEFINE_ORDERTYPE_SIEMENS As String = "SIEMS"               'シーメンス　シーメンス亀田医療情報システム

        Public Const ORDER_DEFINE_UPDATE_USER As String = "ORDER"   'オーダ連携
        Public Const ORDER_DEFINE_UPDATE_KAMEDA_USER As String = "KAMEDA_ORDER"   'オーダ連携

        Public Const AUTO_EXEC_USER As String = "AUTOPG"            '自動PG

        Public Const ORDER_TELMDIV_HEADER As String = "hd"           'ヘッダー部分
        Public Const ORDER_TELMDIV_DATA As String = "data"           'データ部分

        Public Const OrderRfctDiv_NotRfct As String = "0"           '未処理
        Public Const OrderRfctDiv_Rfct As String = "1"              '反映済
        Public Const OrderRfctDiv_RfctError As String = "9"         '反映不可
        Public Const OrderRfctDiv_Cancel As String = "X"            '取り消し
        Public Const OrderRfctDiv_Hold As String = "H"              '保留

        Public Const OrderChangeDivNm_ComExecDiv As String = "ComExecDiv"   '共通処理区分
        Public Const OrderChangeDivNm_ComOdDiv As String = "ComOdDiv"       '共通更新区分

        Public Enum enmOrderPpIsDiv
            NotPrt = 0          '未発行
            Prt = 1             '発行済
            NoPpIs = 9          '対象外
        End Enum

        Public Const OrderAtvOd_Ji As String = "1"              '実施オーダ
        Public Const OrderAtvOd_Oth As String = ""              'それ以外

        Public Enum enmOrderExecDiv
            Insert = 1          '新規追加
            Update = 2          '修正
            Delete = 3          '削除
        End Enum

        Public Enum enmOrderNoFrm
            ShowFrm = 0             '画面表示
            NotFrm = 1          '画面非表示
        End Enum


        '共通指示区分
        Public Const OrderComOdDiv_InHospital As String = "INN"                 '入院
        Public Const OrderComOdDiv_OutHospital As String = "OUT"                '退院
        Public Const OrderComOdDiv_CancelInHospital As String = "DEI"           '入院取消
        Public Const OrderComOdDiv_CancelOutHospital As String = "DEO"          '退院取消
        Public Const OrderComOdDiv_ChangeInHospital As String = "INC"           '入院日変更
        Public Const OrderComOdDiv_ChangeOutHospital As String = "OTC"          '退院日変更
        Public Const OrderComOdDiv_CLINIC_DIALYSIS As String = "CLI"            '外来・透析

        Public Const OrderComOdDiv_UnLimitInHospital As String = "UIN"          '無条件入院（入院中でもエラーとしない）

        Public Const OrderComOdDiv_NoMeal As String = "NOM"                     '欠食
        Public Const OrderComOdDiv_NoMealInHospital As String = "NIN"           '欠食入院
        Public Const OrderComOdDiv_NoMealMove As String = "NMV"                 '欠食移動

        Public Const OrderComOdDiv_GoOut As String = "GOT"                      '外出
        Public Const OrderComOdDiv_SleepOut As String = "SLP"                   '外泊
        Public Const OrderComOdDiv_ComeBack As String = "CBK"                   '帰院
        Public Const OrderComOdDiv_ComeBackEat As String = "CBE"                '帰院+食事変更

        Public Const OrderComOdDiv_AllDataChange As String = "ALL"          '全項目変更

        Public Const OrderComOdDiv_EatChange As String = "EAT"              '食事変更
        Public Const OrderComOdDiv_EatChangeRepeat As String = "EAR"        '食事変更(連続)

        Public Const OrderComOdDiv_SelectEat As String = "SEL"              '選択食

        Public Const OrderComOdDiv_Move As String = "MOV"                   '移動情報更新
        Public Const OrderComOdDiv_WardChange As String = "WRD"             '転棟
        Public Const OrderComOdDiv_RoomChange As String = "ROM"             '転室
        Public Const OrderComOdDiv_MedicalChange As String = "MCA"          '転科

        Public Const OrderComOdDiv_WardMedicalChange As String = "WMR"      '転棟+転科
        Public Const OrderComOdDiv_RoomMedicalChange As String = "RMM"      '転室+転科

        Public Const OrderComOdDiv_DeliveryChange As String = "MDE"         '配膳先変更
        Public Const OrderComOdDiv_DeliveryAddExtensionMeal As String = "EML"      '延食


        Public Const OrderComOdDiv_PersonalInfo As String = "INF"           '個人情報(マスタ＋INF＋ILL)
        Public Const OrderComOdDiv_ChangeDoctor As String = "DCT"           '担当者変換

        Public Const OrderComOdDiv_TPerM As String = "PMS"                  '個人マスタ
        Public Const OrderComOdDiv_TPerEat As String = "PET"                '個人食事情報
        Public Const OrderComOdDiv_TPerMov As String = "PMV"                '個人移動情報
        Public Const OrderComOdDiv_TPerInf As String = "PIF"                '個人情報
        Public Const OrderComOdDiv_TPerDlv As String = "PDV"                '個人配膳情報
        Public Const OrderComOdDiv_TPerCmt As String = "PCM"                '個人コメント情報
        Public Const OrderComOdDiv_TPerLmt As String = "PLT"                '個人制限値情報
        Public Const OrderComOdDiv_TPerIll As String = "PIL"                '個人病名情報

        ''富士通オーダ応答種別
        Public Const OrderEGMAINGX_Ans_OK As String = "OK"                  'OK
        Public Const OrderEGMAINGX_Ans_NG As String = "NG"                  'NG（接続異常）
        Public Const OrderEGMAINGX_Ans_N1 As String = "N1"                  'N1（リトライ）
        Public Const OrderEGMAINGX_Ans_N2 As String = "N2"                  'N2（スキップ）
        Public Const OrderEGMAINGX_Ans_N3 As String = "N3"                  'N3（ダウン）
        Public Const OrderEGMAINGX_Ans_N4 As String = "N4"                  'N4（データなし）
        Public Const OrderEGMAINGX_Ans_N5 As String = "N5"                  'N5（データなし）
        Public Const OrderEGMAINGX_Ans_N6 As String = "N6"                  'N6（更新エラー）

        Public Const OrderEGMAINGX_TelMDiv_ME As String = "ME"                  'ME
        Public Const OrderEGMAINGX_TelMDiv_NI As String = "NI"                  'NI
        Public Const OrderEGMAINGX_TelMDiv_NS As String = "NS"                  'NS
        Public Const OrderEGMAINGX_TelMDiv_KJ As String = "KJ"                  'KJ
        Public Const OrderEGMAINGX_TelMDiv_CE As String = "CE"                  'CE(栄養連携用)

        '変換処理
        Public Const OCHG_OrderMelCnvStr As String = "     "                    '食種コード空白5ケタ
        Public Const OCHG_OrderMCkCnvStr As String = "     "                    '主食コード空白5ケタ
        Public Const OCHG_OrderAgeCdCnvStr As String = "   "                    '年齢性別コード3ケタ
        Public Const OCHG_OrderWardCnvStr As String = "     "                   '病棟コード5ケタ
        Public Const OCHG_OrderRoomCnvStr As String = "     "                   '病室コード5ケタ
        Public Const OCHG_OrderDelivCntStr As String = "     "                  '配膳先コード5ケタ
        Public Const OCHG_OrderCmtCnvStr As String = "        "                 'コメントコード8ケタ
        Public Const OCHG_OrderEatTmCnvInt As Integer = 0                       '時間帯　指定なし

        Public Const OCHG_OrderMealCnvSet As String = "MEAL"                    '食種コード
        Public Const OCHG_OrderMaCkCnvSet As String = "MACK"                    '主食コード
        Public Const OCHG_OrderAgeCnvSet As String = "AGE"                      '年齢コード
        Public Const OCHG_OrderEatTmCnvSet As String = "EATTMDIV"               '時間帯
        Public Const OCHG_OrderCmt1CnvSet As String = "CMT1"                    'コメント１
        Public Const OCHG_OrderCmt2CnvSet As String = "CMT2"                    'コメント２
        Public Const OCHG_OrderCmt3CnvSet As String = "CMT3"                    'コメント３
        Public Const OCHG_OrderWardCnvSet As String = "WARD"                    '病棟コード

        ''PHRMAKER応答種別
        Public Const OrderPHRMAKER_TelMDiv_PMS As String = "PMS"                  '患者情報
        Public Const OrderPHRMAKER_TelMDiv_INN As String = "UIN"                  '無条件入院情報
        Public Const OrderPHRMAKER_TelMDiv_OUT As String = "OUT"                  '退院情報
        Public Const OrderPHRMAKER_TelMDiv_EAT As String = "EAT"                  '食事情報
        Public Const OrderPHRMAKER_TelMDiv_MOV As String = "MOV"                  '移動情報
        Public Const OrderPHRMAKER_TelMDiv_SCK As String = "SCK"                  '病名情報
        Public Const OrderPHRMAKER_TelMDiv_CMT As String = "CMT"                  'コメント情報
        Public Const OrderPHRMAKER_TelMDiv_DAY As String = "DAY"                  '通院情報
        Public Const OrderPHRMAKER_TelMDiv_NOM As String = "NOM"                  '絶食情報

        'アレルギー区分
        Public Const OrderRsheetNm_Allergy As String = "ALLERGY"                'アレルギー患者

        '削除の復活更新年月日
        Public Const OrderDelYMDUpdDiv_NowDelYmd As String = "1"                '削除年月日で更新

        ''SBSオーダ種別
        Public Const OrderSBS_TelMDiv_SL As String = "SL"                           'SL
        Public Const OrderSBS_TelMDiv_TB As String = "TB"                           'TB
        Public Const OrderSBS_TelMDiv_WH As String = "WH"                           'WH
        Public Const OrderSBS_TelMDiv_OD As String = "OD"                           'OD

        Public Const OrderSBS_TelMDiv_UPD As String = "UPD"                         'UPD
        Public Const OrderSBS_TelMDiv_UPWH As String = "UPWH"                       'UPWH

        ''SOFTSオーダ種別
        Public Const OrderSOFT_TelMDiv_CSV As String = "CSV"                        'CSV

        ''ナイスオーダ種別
        Public Const OrderNAIS_TelMDiv_P As String = "P"    '患者情報レコード
        Public Const OrderNAIS_TelMDiv_K As String = "K"    '患者注意レコード
        Public Const OrderNAIS_TelMDiv_T As String = "T"    '給食期間レコード
        Public Const OrderNAIS_TelMDiv_D As String = "D"    '給食詳細レコード

        ''アリスオーダ種別
        Public Const OrderAlice_TelMDiv_TEXT As String = "TEXT"                        'TEXT

        ''MICオーダ種別
        Public Const OrderMIC_TelMDiv_TEXT As String = "TEXT"                        'TEXT

        ''SIEMENSオーダ種別
        Public Const OrderSIEMS_TelMDiv_HEAD As String = "HEAD"                        'ヘッダー部分
        Public Const OrderSIEMS_TelMDiv_TEXT As String = "TEXT"                        'ヘッダーを除くテキスト部分

    End Class

#End Region

    '**********************************************************************
    'その他
    '**********************************************************************
    Public Const DelimiterChar As Char = ":"c          'コンボリスト等でのコード＆名称の区切り文字
    Public Const REG_AJIDAS7 As String = "SOFTWARE\FUJITELECOM\AjiDAS7"     'AjiDAS7のレジストリキー
    Public Const REG_NAME_L As String = "L"     'ライセンス認証のレジストリ名
End Class
