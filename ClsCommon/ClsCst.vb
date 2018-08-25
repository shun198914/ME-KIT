Option Strict On
'**************************************************************************
' Name       :ClsCst ���ʒ萔
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
    '�R���X�^���g�萔
    '**********************************************************************
    '**********************************************************************
    '�t�H���_��
    '**********************************************************************
    Public Class FldrName
        Public Const AppRoot As String = "%AppRoot%"               '�A�v���P�[�V�����S�̂̃��C���t�H���_�ے�����(FrmMenu.exe�̂���p�X)
        Public Const Image As String = "Image"                     '�摜�̓����Ă���t�H���_
        Public Const Sound As String = "Sound"                     '�T�E���h�̓����Ă���t�H���_
        Public Const Bin As String = "Bin"                         '���s�t�@�C���̓����Ă���t�H���_
        Public Const Log As String = "Log"                         '���O�t�@�C���̓����Ă���t�H���_
        Public Const Order As String = "Order"                     '�I�[�_�A�g�p�t�@�C���̓����Ă���t�H���_
        Public Const SqlScript As String = "SqlScript"             'SQL�X�N���v�g�t�@�C���̓����Ă���t�H���_
        Public Const Help As String = "Help"                       '�w���v�t�@�C���̓����Ă���t�H���_
        Public Const Receive As String = "Receive"                 '�_�E�����[�h����t�@�C����{�t�H���_(�ȉ��̃T�u�t�H���_�̐e�t�H���_)
        Public Const ReceiveMst As String = Receive & "\Mst"       '�}�X�^�f�[�^���_�E�����[�h����t�H���_
        Public Const ReceiveApl As String = Receive & "\Apl"       '�A�v���P�[�V�������_�E�����[�h����t�H���_
        Public Const ReceiveVpf As String = Receive & "\Vpf"       '�E�B���X�p�^�[���t�@�C�����_�E�����[�h����t�H���_
        Public Const TempExp As String = "TempExp"                 '���k�t�@�C�����𓀂���e���|�����t�H���_
        Public Const TempExpMst As String = TempExp & "\Mst"       '���k�}�X�^�t�@�C�����𓀂���e���|�����t�H���_
        Public Const TempExpApl As String = TempExp & "\Apl"       '���k�A�v���P�[�V�������𓀂���e���|�����t�H���_
        Public Const TempExpVpf As String = TempExp & "\Vpf"       '���k�E�B���X�p�^�[���t�@�C�����𓀂���e���|�����t�H���_
        Public Const ZipBack As String = "ZipBack"                 '���k�t�@�C����ۊǂ���t�H���_
        Public Const ZipBackMst As String = ZipBack & "\Mst"       '���k�}�X�^�t�@�C����ۊǂ���t�H���_
        Public Const ZipBackApl As String = ZipBack & "\Apl"       '���k�A�v���P�[�V������ۊǂ���t�H���_
        Public Const ZipBackVpf As String = ZipBack & "\Vpf"       '���k�E�B���X�p�^�[���t�@�C����ۊǂ���t�H���_
    End Class

    '**********************************************************************
    '�t�@�C����
    '**********************************************************************
    Public Class FileName
        Public Const ConfigClient As String = "ConfigC.xml"        'XML�N���C�A���g�ݒ�t�@�C��
        Public Const ConfigBackUp As String = "ConfigM.xml"        'XML�o�b�N�A�b�v�ݒ�t�@�C��
        Public Const ConfigServer As String = "ConfigS.xml"        'XML�T�[�o�[�ݒ�t�@�C��
        Public Const LogClient As String = "LogC.txt"              '�N���C�A���g���O
        Public Const LogServer As String = "LogS.txt"              '�T�[�o�[���O
        Public Const LogAct As String = "ActLog.txt"               '�ғ����O
    End Class

    '**********************************************************************
    '�\���ʖ�
    '**********************************************************************
    Public Class FormName
        Public Const FrmMenu As String = "FrmMenu"                  '��ɕ\������郁�j���[���(����FrmMenu.exe�̃v���Z�X��)
        Public Const FrmLogIn As String = "FrmLogIn"                '���O�C�����
        Public Const FrmMain As String = "FrmMain"                  '�o�b�N�O�����h���
        Public Const FrmStartLogo As String = "FrmStartLogo"        '�X�^�[�g���S���
        Public Const FrmMstFd As String = "FrmMstFd"                '�H�i�}�X�^
        Public Const FrmMstFdR As String = "FrmMstFdR"              '�H�i�}�X�^�A�[
        Public Const FrmMstCok As String = "FrmMstCok"              '�����}�X�^
        Public Const FrmMstCokR As String = "FrmMstCokR"            '�����}�X�^�A�[
        Public Const FrmFoodSearch As String = "FrmFoodSearch"      '�H�i����
        Public Const FrmCookSearch As String = "FrmCookSearch"      '��������
        Public Const FrmHaseiTen As String = "FrmHaseiTen"          '�����_��ʁi�ޗ��W�v�j
        Public Const FrmZaiOut As String = "FrmZaiOut"              '�݌ɕ��o��ʁi�P���P�ʁj
        Public Const FrmZaiOutTmDiv As String = "FrmZaiOutTmDiv"        '�݌ɕ��o��ʁi���ԋ敪�j
        Public Const FrmPerChangeHospi As String = "FrmPerChangeHospi"  '���@���ύX���
        Public Const FrmZaikoNoDiv As String = "FrmZaikoNoDiv"          '�I���݌ɉ��(�敪�Ȃ�)
        Public Const FrmChgJiCd As String = "FrmChgJiCd"                    '���Ə��R�[�h�ύX���
        Public Const FrmClearTable As String = "FrmClearTable"            '�V�K���Ə��p�f�[�^�ꊇ�폜���
        '�����n��ʖ�
        Public Shared HachuForms As String() = New String() {"FrmDlveChange", "FrmHachuEdit", "FrmHachuFood", "FrmHachuHistory" _
                                                           , "FrmHachuList", "FrmHachuSend", "FrmHachuSendOK", "FrmMaterials"}
        '2015/10/06 hirabayashi �A�C�R�[���f�B�J�� �����V�X�e�����C
        Public Const FrmMTraderMD As String = "FrmMTraderMD"            '�[�i�j���ݒ���
    End Class

    '**********************************************************************
    '�w�l�k�Z�N�V������
    '**********************************************************************
    Public Class xmlSectionName
        Public Const DBSection As String = "DBSection"             'DB�Z�N�V����

        '*** �� *** �V�X�e���ɂĖ��g�p�ׁ̈A�R�����g�A�E�g Update by K.Aoki *******
        'Public Const ExeToExeRmt As String = "exetoexe.remoting"   'Exe�ԒʐM�p�Z�N�V����
        '*** �� *** �V�X�e���ɂĖ��g�p�ׁ̈A�R�����g�A�E�g Update by K.Aoki *******
    End Class

    '**********************************************************************
    '�w�l�k�Z�N�V�����^�O��
    '**********************************************************************
    Public Class xmlTagName
        Public Const InitialCatalog As String = "InitialCatalog"   '�C�j�V�����J�^���O
        Public Const DataSource As String = "DataSource"           '�f�[�^�\�[�X
        Public Const User As String = "User"                       '���[�U
        Public Const Password As String = "Password"               '�p�X���[�h
        Public Const ConnectTimeout As String = "ConnectTimeout"   '�ڑ��^�C���A�E�g
        Public Const CommandTimeout As String = "CommandTimeout"   '�R�}���h�^�C���A�E�g
        Public Const Pooling As String = "Pooling"                 '�ڑ��v�[���ݒ�
    End Class

    '**********************************************************************
    '�f�[�^�x�[�X�I�u�W�F�N�g��
    '**********************************************************************
    Public Class DataObjName
        Public Const Table As String = "Table"
    End Class

    '**********************************************************************
    '�����O���[�v�i�����O���[�v�̂݁j
    '**********************************************************************
    Public Class Kgn
        Public Const System As String = "System"                   '�V�X�e���J����
    End Class

    '**********************************************************************
    '����
    '**********************************************************************
    Public Class Length
        Public Const CkNm As Integer = 30                          '������
    End Class

    '**********************************************************************
    '�d�m�t�l�萔
    '**********************************************************************
    '���O�^�C�v
    Public Enum enmLogType
        Information = 1
        Warning = 2
        [Error] = 3
        MethodST = 4
        MethodED = 5
        MsgBox = 6
    End Enum

    '���b�Z�[�W�{�b�N�X�A�C�R��
    Public Enum enmMsgBoxIcon
        None = 0
        Information = 1
        Warning = 2
        [Error] = 3
        Question = 4
    End Enum

    '**********************************************************************
    '�d�m�t�l�萔
    '**********************************************************************
    '�l���O�^�C�v
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
    '�e�[�u���̃R���X�g�l
    '**************************************************************************
    Public Class TableConst

        Public Const TBL_FdADiv_Normal As Integer = 0           ''�w��H�iA�敪 0=���
        Public Const TBL_FdADiv_Tube As Integer = 1             ''�w��H�iA�敪 1=�����H
        Public Const TBL_FdADiv_Event As Integer = 2            ''�w��H�iA�敪 2=�s���H
        Public Const TBL_FdADiv_TTP As Integer = 3              ''�w��H�iA�敪 3=TTP

        Public Const TBL_StkDiv_Fresh As Integer = 0            ''�݌Ɏw��敪 0=���N
        Public Const TBL_StkDiv_Zai As Integer = 1              ''�݌Ɏw��敪 1=�݌�
        Public Const TBL_StkDiv_NotFood As Integer = 2          ''�݌Ɏw��敪 2=�H�i�O
        Public Const TBL_StkDiv_NotHac As Integer = 3           ''�݌Ɏw��敪 3=�����ΏۊO

        Public Const TBL_HRound_0 As Integer = 0                ''�������̊ۂ� 0=����
        Public Const TBL_HRound_1 As Integer = 1                ''�������̊ۂ� 1=��1��
        Public Const TBL_HRound_2 As Integer = 2                ''�������̊ۂ� 2=��2��

        Public Const TBL_HCalDiv_RdUp As Integer = 1            ''�������v�Z�敪 1=�؂�グ
        Public Const TBL_HCalDiv_RdDown As Integer = 2          ''�������v�Z�敪 2=�؎̂�
        Public Const TBL_HCalDiv_Round As Integer = 3           ''�������v�Z�敪 3=�l�̌ܓ�

        Public Const TBL_HCostLv_1 As Integer = 1               ''�P�����x�� 1=���x��1
        Public Const TBL_HCostLv_2 As Integer = 2               ''�P�����x�� 2=���x��2

        Public Const TBL_HSalLv_1 As Integer = 1                ''�����̔����x�� 1=���x��1
        Public Const TBL_HSalLv_2 As Integer = 2                ''�����̔����x�� 2=���x��2

        Public Const TBL_NumOfDiv_None As Integer = 0           ''���萔�v�Z�敪 0=�Ȃ�
        Public Const TBL_NumOfDiv_Avg As Integer = 1            ''���萔�v�Z�敪 1=����
        Public Const TBL_NumOfDiv_Max As Integer = 2            ''���萔�v�Z�敪 2=�ő�
        Public Const TBL_NumOfDiv_Min As Integer = 3            ''���萔�v�Z�敪 3=�ŏ�

        Public Const TBL_EdSalDiv_None As Integer = 0           ''�I���敪 0=�I���Ȃ�
        Public Const TBL_EdSalDiv_End As Integer = 1            ''�I���敪 1=�I��

        Public Const TBL_TempZn_Hot As String = "D"             ''���x�ы敪 D=�퉷
        Public Const TBL_TempZn_Cold As String = "C"            ''���x�ы敪 C=�①
        Public Const TBL_TempZn_Ice As String = "F"             ''���x�ы敪 F=�Ⓚ

        Public Const TBL_ZRptDiv_None As Integer = 0            ''�I���񍐋敪 0=�񍐂Ȃ�
        Public Const TBL_ZRptDiv_Rept As Integer = 1            ''�I���񍐋敪 1=��

        Public Const TBL_CompDiv_Nor As Integer = 0             ''�H�i�\���敪 0=�ʏ�
        Public Const TBL_CompDiv_Comp As Integer = 1            ''�H�i�\���敪 1=�\���H�i
        Public Const TBL_CompDiv_Parts As Integer = 2           ''�H�i�\���敪 2=�p�[�c

        Public Const TBL_CompCalDiv_None As Integer = 0         ''�H�i�\���v�Z�敪 0=�v�Z���Ȃ�
        Public Const TBL_CompCalDiv_Cal As Integer = 1          ''�H�i�\���v�Z�敪 1=�v�Z����

        Public Const TBL_NoSelDiv_Sel As Integer = 0            ''�������O�敪 0=�ʏ�
        Public Const TBL_NoSelDiv_NoSel As Integer = 1          ''�������O�敪 1=�������O

        Public Const TBL_Auth_Jigyo As Integer = 0              ''���� 0=�ʏ�
        Public Const TBL_Auth_Siten As Integer = 10             ''���� 10=�ʏ�
        Public Const TBL_Auth_Honsya As Integer = 20            ''���� 20=�ʏ�
        Public Const TBL_Auth_Kikan As Integer = 30             ''���� 30=�ʏ�

        Public Const TBL_SKey_None As Integer = 0               ''�ی�L�[�i���ʁj 0=�ی삵�Ȃ�
        Public Const TBL_SKey_Protected As Integer = 1          ''�ی�L�[�i���ʁj 1=�ʏ�

        Public Const TBL_KonDiv_Kondate As Integer = 0          ''�����敪  0=�ʏ팣��
        Public Const TBL_KonDiv_Per As Integer = 1              ''�����敪  1=�l����

        Public Const TBL_CkTtpDiv_Normal As Integer = 0         ''����TTP�敪  0=�ʏ�
        Public Const TBL_CkTtpDiv_Ttp As Integer = 1            ''����TTP�敪  1=TTP

    End Class

    '�H������
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

    '�H�i�\���敪
    Public Enum enmCompDiv
        Normal = 0    '0:�ʏ�H�i
        Kousei = 1    '1:�H�i�\��
        Parts = 2     '2:�p�[�c�\��
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

    '�����敪
    Public Enum enmNtDiv
        Standard = 1 '��{����
        Amino = 2    '�A�~�m�_
        Fatacid = 3  '���b�_
        User = 99    '���[�U�[
    End Enum

    '�[�������ʒu
    Public Enum enmRoundPosition
        P1 = 1 '�����_�ȉ�1����[�������ΏۂƂ��܂��B(���ʒl�͏����������܂���B��:1)
        P2 = 2 '�����_�ȉ�2����[�������ΏۂƂ��܂��B(���ʒl�͏������ő�1�������܂��B��:1.3)
        P3 = 3 '�����_�ȉ�3����[�������ΏۂƂ��܂��B(���ʒl�͏������ő�2�������܂��B��:1.31)
        P4 = 4 '�����_�ȉ�4����[�������ΏۂƂ��܂��B(���ʒl�͏������ő�3�������܂��B��:1.312)
        P5 = 5 '�����_�ȉ�5����[�������ΏۂƂ��܂��B(���ʒl�͏������ő�4�������܂��B��:1.3123)
        '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
        P15 = 15 '�����_�ȉ�15����[�������ΏۂƂ��܂��B(���ʒl�͏������ő�14�������܂��B)
        '����Add 20110304 FMS)Y.Kurosu Decimal�^�v�Z�̒[�������Ή�
    End Enum

    '�[�������敪
    Public Enum enmRoundDiv
        RoundUp = 1 '�؂�グ
        CutOff = 2  '�؂�̂�
        Rounded = 3 '�l�̌ܓ�
    End Enum

    '�[�������敪2
    Public Enum enmCutUpDiv
        RoundUp = 1 '�؂�グ
        CutOff = 2  '�؂�̂�
    End Enum

    '�}�X�^�e�[�u��
    Public Enum enmMstTables
        SFrm = 1      '��ʏ��
        SSys = 2      '�V�X�e�����
        SMenu = 3     '���j���[���
        MUser = 4     '���[�U�[���
        MName = 5     '���̃}�X�^
        MMealTpP = 6  '�l�H��}�X�^
        MMealTpK = 7  '�����H��}�X�^
        MTrader = 8   '�d����}�X�^
        MNamCls = 9   '���̕��ރ}�X�^
        MMealGpP = 10 '�l�H��O���[�v�}�X�^
        MMealGpK = 11 '�����H��O���[�v�}�X�^
        MGpFd = 12    '�H�i�Q�}�X�^(1or2)
        MSisetsu = 13 '�{�݃}�X�^
        MGpFdCls = 14 '�H�i�Q���ރ}�X�^(1or2)
        MKMealTp = 15 '�����}�X�^�p�H��
        MHeiSt = 16   '���݃}�X�^

        '����������������������������������������
        '�� �ǂ�ǂ�ǉ����ĉ������B�i��ͥ)�@��
        '����������������������������������������

    End Enum

    '����_�������f�[�^�敪
    Public Enum enmMNameNtDataDiv
        CdNm = 1   '��������
        CdNmAb = 2 '����
        CdCr = 3   '�P��
        CdDic = 4  '�����_����
        SortNo = 5 '�\����
    End Enum

    '��V�X�e���^�f�s�w���ʃt���O
    Public Enum enmSystem
        K_Sys = 1   '��V�X�e��
        GTX = 2     '�f�s�w
    End Enum

    '�L���t���O
    Public Enum enmValidF
        None = 0    '����
        Exists = 1  '�L��
    End Enum

    '�P�ʌv�Z�敪
    Public Enum enmCrCalDiv
        Normal = 0    '�ʏ�v�Z
        Unit = 1      '�P�ʂŌv�Z�i�����V�X�e����F�j
    End Enum

    '**************************************************************************
    '�����}�X�^
    '**************************************************************************
    Public Enum enmCkDiv
        Normal = 0  '�ʏ�
        MainCok = 1 '��H
        Tube = 2    '�o�ǉh�{
        Milk = 3    '����   
        Vena = 4    '�Ö��h�{
    End Enum


    '**************************************************************************
    '�H�i�}�X�^
    '**************************************************************************
    '-------------------------------------
    '�݌Ɏw��敪
    '-------------------------------------
    Public Enum enmStkDiv
        Fresh = 0       '���N
        Stock = 1       '�݌�
        NotFood = 2     '�H�i�O
        NotHac = 3      '�����ΏۊO
    End Enum

    '-------------------------------------
    '�������̊ۂ�
    '-------------------------------------
    Public Enum enmHRound
        AMOUNT_0 = 0    '����
        AMOUNT_1 = 1    '�����_��P��
        AMOUNT_2 = 2    '�����_��Q��
    End Enum

    '-------------------------------------
    '�������v�Z�敪
    '-------------------------------------
    Public Enum enmHCalDiv
        RoundDUp = 1    '�؂�グ
        RoundDown = 2   '�؎̂�
        Round = 3       '�l�̌ܓ�
    End Enum

    '-------------------------------------
    '�P�����x��
    '-------------------------------------
    Public Enum enmHCostLv
        CostLv1 = 1    '���x���P
        CostLv2 = 2    '���x���Q
    End Enum

    '-------------------------------------
    '�����̔����x��
    '-------------------------------------
    Public Enum enmHSalLv
        SalLv1 = 1    '���x���P
        SalLv2 = 2    '���x���Q
    End Enum

    '�� Add by K.Kojima 2008/12/12 Start�i�P�D���M���ɒ����ς̃��R�[�h�ɏI�����܂܂�Ă����ꍇ�̂݌x�����b�Z�[�W��\������l�ɕύX�j*************************************************************************
    '-------------------------------------
    '�I���敪	0:�I���Ȃ��A1:�I���H�i
    '-------------------------------------
    Public Enum enmEdSalDiv
        None = 0        '�I���Ȃ�
        SalEnd = 1      '�I���H�i
    End Enum
    '�� Add by K.Kojima 2008/12/12 End  �i�P�D���M���ɒ����ς̃��R�[�h�ɏI�����܂܂�Ă����ꍇ�̂݌x�����b�Z�[�W��\������l�ɕύX�j*************************************************************************

    '-------------------------------------
    '�I���񍐋敪
    '-------------------------------------
    Public Enum enmZRptDiv
        NotInfo = 0     '�񍐂Ȃ�
        Info = 1        '�񍐂���
    End Enum

    '-------------------------------------
    '����
    '-------------------------------------
    Public Enum enmAuthority
        Kikan = 30      '��V�X�e��
        Honsha = 20     '�{��
        Shiten = 10     '�x�X
        Jigyosyo = 0    '���Ə�
    End Enum


    '-------------------------------------
    '�����敪
    '-------------------------------------
    Public Enum enmNumOfDiv
        NumOfDiv_None = 0 '����
        NumOfDiv_Avg = 1  '����
        NumOfDiv_Max = 2  '�ő�
        NumOfDiv_Min = 3  '�ŏ�
    End Enum

    '-------------------------------------
    '���M�X�e�[�^�X
    '-------------------------------------
    Public Enum enmSendSt
        NoData = -1      '�f�[�^�Ȃ�
        NoSend = 0       '�����M
        NoPrint = 0      '�����
        OneSend = 1      '1��ڑ��M��
        Printded = 1     '����ς�
        TwoSending = 2   '2��ڈȍ~�C����
        TwoSend = 3      '2��ڈȍ~���M��
    End Enum

    '**************************************************************************
    '�H������
    '**************************************************************************
    Public Class EatTime
        Public Const MorningDiv As Integer = 10
        Public Const MorningName As String = "���H"
        Public Const MorningNameShort As String = "��"
        Public Const AM10Div As Integer = 11
        Public Const AM10DivName As String = "���ԐH"
        Public Const AM10DivNameShort As String = "����"
        Public Const LunchDiv As Integer = 20
        Public Const LunchName As String = "���H"
        Public Const LunchNameShort As String = "��"
        Public Const PM15Div As Integer = 21
        Public Const PM15Name As String = "���ԐH"
        Public Const PM15NameShort As String = "����"
        Public Const EveningDiv As Integer = 30
        Public Const EveningName As String = "�[�H"
        Public Const EveningNameShort As String = "�["
        Public Const NightDiv As Integer = 31
        Public Const NightName As String = "�[�ԐH"
        Public Const NightNameShort As String = "�[��"
    End Class

    '-------------------------------------
    '�����i���j
    '-------------------------------------
    Public Const HonAuthNameShort As String = "H"
    Public Const SiAuthNameShort As String = "S"
    Public Const JiAuthNameShort As String = "J"

    '-------------------------------------
    '�戵�I�����i�����������̐ݒ�l = �ڍs�f�[�^�̐ݒ�l�j
    '-------------------------------------

    '�V�Ή� (hayashi)
    Public Const UseEdDtNoLimit As String = "90001231"
    'Public Const UseEdDtNoLimit As String = "20991231"

    '2009/10/28 CM)H.Kuwahara Start 
    '�b�j�r�j�敪
    Public Enum enmCkSkDiv
        Normal = 0      '�ʏ�
        Central = 1     '�Z���g�����L�b�`��
        Satellite = 2   '�T�e���C�g�L�b�`��
    End Enum
    '2009/10/28 CM)H.Kuwahara End 
    Public Const intNutriCnt As Integer = 80

    '-------------------------------------
    '�����敪�i�h�{�w���j
    '-------------------------------------
    Public Class PerExam
        Public Enum enmExamDiv
            ''' <summary>�S��</summary>
            ''' <remarks></remarks>
            All = 0

            ''' <summary>�h�{�w��</summary>
            ''' <remarks></remarks>
            Tec = 1

            ''' <summary>�h�{�Ǘ�</summary>
            ''' <remarks></remarks>
            Mng = 2

            ''' <summary>�h�{�P�A</summary>
            ''' <remarks></remarks>
            Care = 3
        End Enum

        Public Class MNameCd1
            '-------------------------------------
            '�g�̌v�����ڃR�[�h�i�h�{�w���j
            '-------------------------------------
            Public Class Body
                Public Const Height As String = "01010"        '�g��
                Public Const HeightYmd As String = "01020"     '�g�������
                Public Const Hizadaka As String = "01030"      '�G��
                Public Const Weight As String = "01040"        '�̏d
                Public Const WeightYmd As String = "01050"     '�̏d�����
                Public Const BMI As String = "01060"           'BMI
            End Class

            '-------------------------------------
            '�X�N���[�j���O���ڃR�[�h�i�h�{�w���j
            '-------------------------------------
            Public Class Scr
                Public Const PrevWeight As String = "02010"    '�O��̏d
                Public Const PrevWeightYmd As String = "02020" '�O��̏d�����
                Public Const WeightFromTo As String = "02030"  '�̏d�������
                Public Const WeightRatio As String = "02040"   '�̏d�ϓ���
                Public Const Alb As String = "02050"           '�����A���u�~��
                Public Const AlbYmd As String = "02060"        '�����A���u�~�������
                Public Const UseRatio As String = "02070"      '�H���ێ��
                Public Const ScrRefCol As String = "02080"     '�H���ێ���e
                Public Const Keikou As String = "02090"        '�o���h�{
                Public Const Keichou As String = "02100"       '�o���h�{
                Public Const Jyoumyaku As String = "02110"     '�Ö��h�{
                Public Const Jyokusou As String = "02120"      '���
                Public Const Other As String = "02130"         '���̑�
                Public Const Risk As String = "02140"          '���X�N
            End Class

            '-------------------------------------
            '�h�{�⋋�ʎZ�荀�ڃR�[�h�i�h�{�w���j
            '-------------------------------------
            Public Class Nutri
                Public Const Energy As String = "03010"        '�K�v�G�l���M�[
                Public Const Protein As String = "03020"       '�K�v�^���p�N��
                Public Const Fat As String = "03030"           '�K�v���b
                Public Const Carbo As String = "03040"         '�K�v�Y������
                Public Const Water As String = "03050"         '�K�v������
                Public Const Salt As String = "03060"          '�K�v�H����
                Public Const GiveEnergy As String = "03070"    '���^�G�l���M�[
            End Class

            '-------------------------------------
            '�h�{�⋋�ʎZ�荀�ڃR�[�h�i��ʌ������ځj
            '-------------------------------------
            Public Class Stand
                Public Const GOT As String = "04010"           '�f�n�s
                Public Const GPT As String = "04020"           '�f�o�s
                Public Const ALP As String = "04030"           '�`�k�o
            End Class

            '-------------------------------------
            '�h�{�⋋�ʎZ�荀�ڃR�[�h�i�����j
            '-------------------------------------
            Public Class Disease
                Public Const HL As String = "05010"            '��������
                Public Const HLMedic As String = "05020"       '�������ǁE����
                Public Const DM As String = "05030"            '���A�a
                Public Const DMMedic As String = "05040"       '���A�a�E����
                Public Const HT As String = "05050"            '������
                Public Const HTMedic As String = "05060"       '�������E����
                Public Const HD As String = "05070"            '�S����
                Public Const HDMedic As String = "05080"       '�S�����E����
            End Class

            '-------------------------------------
            '�h�{�⋋�ʎZ�荀�ڃR�[�h�i���̑��j
            '-------------------------------------
            Public Class Other
                Public Const Denture As String = "11010"       '�`��
                Public Const PriOpe As String = "11020"        '��p�\��
                Public Const OthRefCol As String = "11030"     '���l
            End Class
        End Class

        '-------------------------------------
        '�������ړ��͕����i�h�{�w���j
        '-------------------------------------
        Public Enum enmNtrExamEditor
            ''' <summary>���l����</summary>
            Dec = 1

            ''' <summary>�e�L�X�g����</summary>
            Txt = 2

            ''' <summary>�R���{���X�g����</summary>
            Cmb = 3

            ''' <summary>���t����</summary>
            Ymd = 4
        End Enum
    End Class

    '**************************************************************************
    '����
    '**************************************************************************
    '**************************************************************************
    '��ʋ��ʒ�`
    '**************************************************************************
    '**********************************************************************
    '�d����}�X�^
    '**********************************************************************
    Public Class MTraderSec
        Public Const TBL_DEFINE_DLVE As String = "0,����,-1,�O��,-2,�O�X��" '�[�i��
    End Class
    '**********************************************************************
    '���̃}�X�^
    '**********************************************************************
    Public Class MNameDisCd
        Public Const ElementStandard As String = "0001"             '��{������
        Public Const ElementAmino As String = "0002"                '�A�~�m�_������
        Public Const ElementFatacid As String = "0003"              '���b�_������
        Public Const ExDiadetesDiv As String = "0004"               '���A�����\�敪    
        Public Const ExKidneyDiv As String = "0005"                 '�t�������\�敪
        Public Const ExDiaKidneyDiv As String = "0006"              '���A�a�t���ǌ����\�敪
        Public Const Credit As String = "0010"                      '�P��
        Public Const SerchClass As String = "0011"                  '��������
        Public Const AppFoodDiv As String = "0012"                  '�w��H�i�敪
        Public Const MethodThawing As String = "0030"               '�𓀕��@
        Public Const MethodCut As String = "0033"                   '�J�b�g���@
        Public Const MethodLowproc As String = "0034"               '���������@
        Public Const CookType As String = "0035"                    '�������
        Public Const MaterialsType As String = "0036"               '�ޗ����
        Public Const CookingForm As String = "0037"                 '�����`��
        Public Const Plate As String = "0038"                       '�����M
        Public Const MealtypeSumPtan1 As String = "0050"            '�H��W�v�p�^�[���P
        Public Const MealtypeSumPtan2 As String = "0051"            '�H��W�v�p�^�[���Q
        Public Const MealtypeSumPtan3 As String = "0052"            '�H��W�v�p�^�[���R
        Public Const MealtypeSumPtan4 As String = "0053"            '�H��W�v�p�^�[���S
        Public Const SddishesForm As String = "0054"                '���H�`��
        Public Const Sum3Div As String = "0055"                     '�W�v�R�敪
        Public Const Sum6Div As String = "0056"                     '�W�v�U�敪
        Public Const StsGpCd As String = "0057"                     '���v�p�H��O���[�v�R�[�h
        ' �� 2011/06/17 INSERT N.Mimori
        Public Const CommentSumClass As String = "0058"             '�R�����g�`��
        ' �� 2011/06/17 INSERT N.Mimori
        Public Const KonForm As String = "0059"                     '�����p�`��
        Public Const Tray As String = "0060"                        '�g���[
        Public Const MaCkClsCd As String = "0061"                   '��H����

        '��Add by K.Ikeda 2011/02/02 (�P�D���̃}�X�^�e�[�u���̕ύX�ɑΉ�)**************************************************
        Public Const MealtypeSumPtan5 As String = "0065"            '�H��W�v�p�^�[���T
        Public Const MealtypeSumPtan6 As String = "0066"            '�H��W�v�p�^�[���U
        Public Const MealtypeSumPtan7 As String = "0067"            '�H��W�v�p�^�[���V
        Public Const MealtypeSumPtan8 As String = "0068"            '�H��W�v�p�^�[���W
        '��Add by K.Ikeda 2011/02/02 (�P�D���̃}�X�^�e�[�u���̕ύX�ɑΉ�)**************************************************

        Public Const Ward As String = "2001"                        '�a��
        Public Const SickRoom As String = "2002"                    '�a��
        Public Const Consulting As String = "2003"                  '�f�É�
        Public Const LifeCoefficient As String = "2004"             '�g�̊������x��
        Public Const UnderReason As String = "2005"                 '���H���R
        Public Const DeliverPlace As String = "2006"                '�z�V����
        Public Const SickName As String = "2007"                    '�a��
        Public Const BabyBottle As String = "2008"                  '�M���r
        Public Const Nipple As String = "2009"                      '����
        Public Const Byoutai As String = "2010"                     '�a�ԐH
        Public Const WardClsCd As String = "2011"                   '�a������
        Public Const InputPlate As String = "2012"                  '�l�ێ旦���͎M

        '��Add by S.Shin 2012/10/18 **************************************************
        Public Const NutMealMethod As String = "2013"               '���^���@
        Public Const NutMealCnt As String = "2014"                  '���^��
        Public Const NutOneVal As String = "2015"                   '1���
        '��Add by S.Shin 2012/10/18 **************************************************

        Public Const MaterialsSumDiv As String = "3001"             '�ޗ��W�v�敪
        Public Const JournalizingDiv As String = "3002"             '�d��敪
        Public Const Warehouse As String = "4001"                   '�q�ɋ敪

        '��Add by Hirata 2015/10/07 **************************************************
        Public Const FacDiv As String = "4010"               '�H��敪
        '��Add by Hirata 2015/10/07 **************************************************

        Public Const DistrictName As String = "5001"                '�n�於
        Public Const CkClDay As String = "5002"                     '�N�b�N�`�����t
        Public Const VcCkDay As String = "5003"                     '�^�󒲗����t
        '20111215 S
        Public Const CtlTm As String = "6001"                       '�\�񎞊ԑ�
        Public Const Dietician As String = "6002"                   '�h�{�m
        Public Const Resv As String = "6003"                        '�w���\��g
        Public Const Tec As String = "6004"                         '�w���敪
        Public Const Stress As String = "6005"                      '�X�g���X�W��
        Public Const Active As String = "6006"                      '�����W��
        Public Const ExamCtgry As String = "6007"                   '��������
        Public Const ExamItem As String = "6008"                    '��������
        Public Const HeightCalcCareMen As String = "6009"           '�g���Z�o�W��(�h�{�P�A��)(�j)
        Public Const HeightCalcCareWmn As String = "6010"           '�g���Z�o�W��(�h�{�P�A��)(�j)
        Public Const HeightCalcChumMen As String = "6011"           '�g���Z�o�W��(Chumlea��)(�j)
        Public Const HeightCalcChumWmn As String = "6012"           '�g���Z�o�W��(Chumlea��)(��)
        Public Const BeeHarrisMen As String = "6013"                'BEE�W��(�n���X�E�x�l�f�B�N�g��)(�j)
        Public Const BeeHarrisWmn As String = "6014"                'BEE�W��(�n���X�E�x�l�f�B�N�g��)(��)
        Public Const BeeJpnMen As String = "6015"                   'BEE�W��(���{�l�ȈՎ�)(�j)
        Public Const BeeJpnWmn As String = "6016"                   'BEE�W��(���{�l�ȈՎ�)(��)
        Public Const Risk As String = "6017"                        '���X�N�i�R���{�{�b�N�X�p�j
        Public Const PlusMinus As String = "6018"                   '�z���A���i�R���{�{�b�N�X�p�j
        Public Const Exists As String = "6019"                      '�L���i�R���{�{�b�N�X�p�j
        Public Const WeightFromTo As String = "6020"                '�̏d�v������
        '20111215 E
        Public Const Nameofera As String = "9001"                   '�N��
        Public Const SalesTax As String = "9002"                    '�����

        ' �� 2011/06/17 DELETE N.Mimori
        ''7/7���݁AUclSelectFrm.SetItemNameComment�����Ή��ׁ̈A�Ή�������ɍ폜�\��
        ''�� Delete by K.Kojima 2008/07/07 Start�i�P.�R�����g�W�v���ނ͕s�v�ȈׁA�폜�j*************************************************************************
        'Public Const CommentSumClass As String = "0058"             '�R�����g�W�v����
        ''�� Delete by K.Kojima 2008/07/07 End  �i�P.�R�����g�W�v���ނ͕s�v�ȈׁA�폜�j*************************************************************************
        ' �� 2011/06/17 DELETE N.Mimori
    End Class

    '**********************************************************************
    '�l���
    '**********************************************************************
    Public Class PerCST

        '�l���p�R���X�g
        Public Const MINVALUE_T_PerYMD As String = "00000000"           '�l���̓��t�ő�l�i�l�Ȃ��Ӗ��j
        Public Const MINVALUE_T_PerTmDiv As Integer = 0                 '�l���̎��ԋ敪�ő�l�i�l�Ȃ��̈Ӗ��j

        Public Const LIMITVALUE_T_PerYMD As String = "99999999"         '�l���̓��t�ő�l�i�`�����Ƃ̈Ӗ��j
        Public Const LIMITVALUE_T_PerTmDiv As Integer = 99              '�l���̎��ԋ敪�ő�l�i�`�����Ƃ̈Ӗ��j

        '=======================================================
        '2009/01/14 FTC)T.IKEDA �w�i�F��Brown��DarkGray�ɕύX
        '=======================================================
        Public Const NO_MEAL_BACKCOLOR_NAME As String = "DarkGray"
        Public Const NO_MEAL_FORECOLOR_NAME As String = "White"

        Public Const PERCD_LEN As Integer = 10                          '�l�R�[�h����

        '���H�敪
        Public Enum enmNoMeal
            Normal = 0  '�ʏ�
            NoMeal = 1  '���H
        End Enum

        '�l�����敪
        Public Enum enmPerMn
            Normal = 0  '�ʏ팣��
            PerMn = 1   '�l����
        End Enum
        '�l�����ی�敪
        Public Enum enmPerProtect
            Normal = 0      '�Ȃ�
            Protect = 1     '����
        End Enum

        '�������敪
        Public Enum enmKCntDiv
            Normal = 0  '�ʏ�
            Week = 1    '�T��
        End Enum

        '����
        Public Enum enmSex
            Man = 1     '�j
            Woman = 2   '��
        End Enum

        '�폜���O�敪
        Public Enum enmUnDelDiv
            Normal = 0  '�ʏ�
            NoDel = 1   '�폜���O
        End Enum

        '���ʐH���Z�敪
        Public Enum enmEatAdDiv
            NoAdd = 0   '����Z
            Add = 1     '���Z
        End Enum

        '�H�D�敪
        Public Enum enmMCPrt
            None = 0    '�Ȃ�
            Exists = 1  '�L��
        End Enum

        '�w���敪
        Public Enum enmTecDiv
            Personal = 1   '�l
            Group = 2      '�W�c
            Outpatient = 3 '�O��
        End Enum

        'NST�Ή�
        Public Enum enmNstDiv
            NoNst = 0     'NST�Ȃ�
            Nst = 1   'NST�Ή�
        End Enum

        ''���H
        Public Const DELIVERPLACE_EXTENSION_MEAL_CODE As String = "77777"

    End Class

    '�� Add by K.Kojima 2008/12/05 Start�i�P�D�d���敪��\������l�ɕύX�j*************************************************************************
    '**********************************************************************
    '�d�������
    '**********************************************************************
    Public Class BuyCST

        '�d���敪
        Public Enum enmStkDiv
            Fresh = 0       '���N�H�i�i�݌ɊǗ��s�v�j
            Zai = 1         '�݌ɐH�i
            NotFood = 2     '�H�i�O�i�o�������j
            NotHac = 3      '�����ΏۊO
        End Enum

    End Class
    '�� Add by K.Kojima 2008/12/05 End  �i�P�D�d���敪��\������l�ɕύX�j*************************************************************************

    '**********************************************************************
    '�I���}�X�^
    '**********************************************************************
    Public Enum enmZaiDiv
        Patient = 1     '�l
        Other = 2       '�l�O
    End Enum

    '**********************************************************************
    '�����Ɩ�
    '**********************************************************************
    Public Class THacInf

        '�� Change by K.Kojima 2008/09/18 Start�i�P�D�ߋ��̔[�i���̓G���[�A�����[�i���͌x���A�����̔[�i���͐���ƂȂ�l�ɉ��Ή��j*************************************************************************
        Public Const DATA_NO_SEND_DAY As Integer = 0                '�f�[�^�`���s�����i��������f�[�^�`�����ł��Ȃ�������ݒ�j
        'Public Const DATA_NO_SEND_DAY As Integer = 1                '�f�[�^�`���s�����i��������f�[�^�`�����ł��Ȃ�������ݒ�j
        '�� Change by K.Kojima 2008/09/18 End  �i�P�D�ߋ��̔[�i���̓G���[�A�����[�i���͌x���A�����̔[�i���͐���ƂȂ�l�ɉ��Ή��j*************************************************************************

        '�z�M�σt���O
        Public Enum enmSndDiv
            NoDelivery = 0  '���z�M
            Delivery = 1    '�z�M��
        End Enum

        Public Const CST_HAC_DEADLINE_FILE As String = "Aji7HDF"
        Public Const CST_HAC_MEQSSY_FILE As String = "Aji7MekyuHac.csv"

        ''' <summary>
        ''' �H��敪-���n���蕶����(DisCd=4010,Cd1��ZZ�Ŏn�܂���̂͌��n�Ƃ���)
        ''' </summary>
        ''' <remarks>�H��敪-���n���蕶����</remarks>
        Public Const CST_FACDIV_GENTHI_START_CHAR As String = "ZZ"
    End Class

    '���ݎ��Ə��敪
    Public Enum enmHeiStDiv
        Normal = 0      '�ʏ�
        Heisetsu = 1    '����
        Kinrin = 2      '�ߗ�
    End Enum

    '**********************************************************************
    ' �W�z�M�f�[�^�敪
    '**********************************************************************
    Public Class SendRecvDiv
        Public Const SRDiv_Memo_G As String = "GMEM"    '����
    End Class

    '�� Add by K.Kojima 2009/01/06 Start�i�P�D�}�X�^����M���ɓ��e��\������@�\��ǉ��j*************************************************************************
    '**************************************************************************
    '�ő��H�i�Q�}�X�^
    '**************************************************************************
    Public Enum enmGpFdDiv
        First = 1    '��1�H�i�Q
        Second = 2   '��2�H�i�Q
    End Enum

    '**************************************************************************
    '����M�}�X�^�ڍ׉�ʗp
    '**************************************************************************
    Public Class DetailCST

        '**************************************************************************
        '��ʔ��ʕ���
        '**************************************************************************
        Public Const CST_DISTINCTION As String = "FrmMstDetail"

        '**************************************************************************
        '��؂蕶��
        '**************************************************************************
        Public Const CST_DELIMITER As String = "_"

        '**************************************************************************
        '����M���ʕ���
        '**************************************************************************
        Public Class StrSndRcvCST
            Public Const Send As String = "S"           '���M
            Public Const Receive As String = "R"        '��M
        End Class

        '**************************************************************************
        '�}�X�^���ʕ���
        '**************************************************************************
        Public Class StrMstCST
            Public Const MFd As String = "MFd"          '�H�i�}�X�^
            Public Const MNutri As String = "MNutri"    '�����}�X�^
            Public Const MCok As String = "MCok"        '�����}�X�^
        End Class

    End Class
    '�� Add by K.Kojima 2009/01/06 End  �i�P�D�}�X�^����M���ɓ��e��\������@�\��ǉ��j*************************************************************************

    '��hayashi 2013/04/03
    '**************************************************************************
    '�x���_�[�R�[�h
    '**************************************************************************
    Public Class VendorCd
        Public Const NGF As String = "005" '���{�[�l�����t�[�h(NGF)
    End Class

    '**************************************************************************
    '�d����}�X�^�̎w��H�i�敪(NGF�ɂĎg�p)
    '**************************************************************************
    Public Class FdDiv
        Public Const R0 As String = "00000" '�Ȃ�
        Public Const R1 As String = "00001" 'R1
        Public Const R2 As String = "00002" 'R2
        Public Const R3 As String = "00003" 'R3
    End Class
    '��hayashi 2013/04/03

    '**********************************************************************
    ' GtUpdater�@�\
    '**********************************************************************
    Public Class GtUpdater
        '**********************************************************************
        ' SQLCMD�p�p�����[�^
        '**********************************************************************
        Public Const C_SQLSERVER_SCMD As String = "SQLCMD.exe"
        Public Const C_SQLSERVER_USER As String = "-U{0}"
        Public Const C_SQLSERVER_PASSWORD As String = "-P{0}"
        Public Const C_SQLSERVER_DBNAME As String = "-d{0}"
        Public Const C_SQLSERVER_DBSERVER As String = "-S{0}"
        Public Const C_SQLSERVER_COMMAND As String = "-i{0}"

        '�ŐV�ŢGtUpdater��i�[�ꏊ
        Public Const GtUpdaterFD As String = "updater"

        ' �f�[�^�J�^���O���񋓌^
        Public Enum enumDtClg
            DBServer = 0        ' DBServerName
            DBName = 1          ' DBName
            UserID = 2          ' UserID
            Password = 3        ' Password
            Max = 3             ' �ő�C���f�N�X�ԍ�
        End Enum

    End Class

    '**********************************************************************
    ' �o�b�N�A�b�v�@�\
    '**********************************************************************
    Public Class Backup

        '**********************************************************************
        '�w�l�k�Z�N�V������
        '**********************************************************************
        Public Class xmlSectionName
            Public Const PathSection As String = "PathSection"         '�p�X�Z�N�V����
            Public Const application As String = "application"         '�A�v���P�[�V�����Z�N�V����
        End Class

        '**********************************************************************
        '�w�l�k�Z�N�V�����^�O��
        '**********************************************************************
        Public Class xmlTagName
            Public Const InitialCatalog As String = "InitialCatalog"    '�C�j�V�����J�^���O
            Public Const DataSource As String = "DataSource"            '�f�[�^�\�[�X
            Public Const User As String = "User"                        '���[�U
            Public Const Password As String = "Password"                '�p�X���[�h
            Public Const ConnectTimeout As String = "ConnectTimeout"    '�ڑ��^�C���A�E�g
            Public Const CommandTimeout As String = "CommandTimeout"    '�R�}���h�^�C���A�E�g
            Public Const Pooling As String = "Pooling"                  '�ڑ��v�[���ݒ�
            Public Const BackupSakiPath As String = "BackupSakiPath"    '�ۑ���t�H���_�����l
            Public Const RestoreMotoPath As String = "RestoreMotoPath"  '�������t�H���_�����l
            Public Const ZipNameSoft As String = "ZipNameSoft"          '�\�t�g���k�t�@�C���������l
            Public Const ZipNameDB As String = "ZipNameDB"              'DB���k�t�@�C���������l
            'Public Const TempPath As String = "TempPath"                '�e���|�����t�H���_�쐬�p�X��
            Public Const Expansion As String = "Expansion"              '���k�t�@�C���g���q��
            Public Const SoftBackName As String = "SoftBackName"        '�\�t�g�o�b�N�A�b�v��
            Public Const DebugFlg As String = "DebugFlg"                '�f�o�b�N���t���O
            Public Const SoftPath As String = "SoftPath"                '�\�t�g���݃t�H���_�p�X
            Public Const ServiceName As String = "ServiceName"          '�T�[�r�X��
            'Public Const DBSavePlace As String = "DBSavePlace"          '�f�[�^�x�[�X�t�@�C���ۑ��ꏊ
            Public Const BakExeFileName As String = "BakExeFileName"    '�o�b�N�A�b�vEXE�t�@�C����
            Public Const BackupFolder As String = "BackupFolder"        '�o�b�N�A�b�v�t�@�C���ۑ��t�H���_
            Public Const CmdLineStr As String = "CmdLineStr"            '�N�������ʃR�}���h���C��������
            Public Const ResExeFileName As String = "ResExeFileName"    '���X�g�AEXE�t�@�C����
        End Class

    End Class

    '**********************************************************************
    '�e�[�u����
    '**********************************************************************
    Public Class TableName
        Public Const M_Fd As String = "M_Fd"                    '�H�i�}�X�^
        Public Const M_FdSKey As String = "M_FdSKey"            '�H�i�ی�L�[�}�X�^
        Public Const M_FdCost As String = "M_FdCost"            '�H�i�P���}�X�^
        Public Const M_FdCmt As String = "M_FdCmt"              '�H�i�R�����g�}�X�^
        Public Const M_FdComp As String = "M_FdComp"            '�H�i�\���}�X�^
        Public Const M_FdNoUse As String = "M_FdNoUse"          '�g�p�����킹�H�i�}�X�^
        Public Const M_Nutri As String = "M_Nutri"              '�����}�X�^
        Public Const M_NutSKey As String = "M_NutSKey"          '�����ی�L�[�}�X�^
        Public Const M_Trader As String = "M_Trader"            '�d����}�X�^
        Public Const M_Cmt As String = "M_Cmt"                  '�R�����g�}�X�^
        Public Const M_CmtChg As String = "M_CmtChg"            '�R�����g�ϊ��}�X�^
        Public Const M_NamCls As String = "M_NamCls"            '���̕��ރ}�X�^
        Public Const M_Name As String = "M_Name"                '���̃}�X�^
        Public Const M_ContCls As String = "M_ContCls"          '�_�񕪗ރ}�X�^
        Public Const M_ContClsLk As String = "M_ContClsLk"      '�_�񕪗ފ֘A�}�X�^
        Public Const M_Cok As String = "M_Cok"                  '�����}�X�^
        Public Const M_CokFd As String = "M_CokFd"              '�����H�i�}�X�^
        Public Const M_MealTpP As String = "M_MealTpP"          '�H��}�X�^�i�l�j
        Public Const M_MealTpK As String = "M_MealTpK"          '�H��}�X�^�i�����j
        Public Const M_MealZai As String = "M_MealZai"          '�H��ދ�}�X�^
        Public Const M_MealLmtP As String = "M_MealLmtP"        '�H�퐧���l�}�X�^(�l�j
        Public Const M_MealLmtK As String = "M_MealLmtK"        '�H�퐧���l�}�X�^(�����j
        Public Const M_MealCmtP As String = "M_MealCmtP"        '�H��R�����g�}�X�^�i�l�j
        Public Const M_MealGpP As String = "M_MealGpP"          '�H��O���[�v�}�X�^�i�l�j
        Public Const M_MealGpK As String = "M_MealGpK"          '�H��O���[�v�}�X�^�i�����j
        Public Const M_GpFdClsMx As String = "M_GpFdClsMx"      '�ő��H�i�Q���ރ}�X�^
        Public Const M_GpFdMx As String = "M_GpFdMx"            '�ő��H�i�Q�}�X�^
        Public Const M_GpFdCls1 As String = "M_GpFdCls1"        '�H�i�Q���ރ}�X�^1
        Public Const M_GpFd1 As String = "M_GpFd1"              '�H�i�Q�}�X�^1
        Public Const M_GpFdLk1 As String = "M_GpFdLk1"          '�H�i�Q�֘A�}�X�^1
        Public Const M_GpFdCls2 As String = "M_GpFdCls2"        '�H�i�Q���ރ}�X�^2
        Public Const M_GpFd2 As String = "M_GpFd2"              '�H�i�Q�}�X�^2
        Public Const M_GpFdLk2 As String = "M_GpFdLk2"          '�H�i�Q�֘A�}�X�^2
        Public Const M_Age As String = "M_Age"                  '�N��Ǘ��}�X�^
        Public Const M_AgeInf As String = "M_AgeInf"            '�N����}�X�^
        Public Const M_AgeDNR As String = "M_AgeDNR"            '�N��v�ʃ}�X�^
        Public Const T_FdGpAvgK As String = "T_FdGpAvgK"        '�H�ƍ\���i���d�j
        Public Const T_FdGpAvgJ As String = "T_FdGpAvgJ"        '�H�ƍ\���i���{�j
        Public Const M_SiTen As String = "M_SiTen"              '�x�X�}�X�^
        Public Const M_JiGyo As String = "M_JiGyo"              '���Ə��}�X�^
        Public Const M_Sisetsu As String = "M_Sisetsu"          '�{�݃}�X�^
        Public Const M_HeiSt As String = "M_HeiSt"              '���݁i�ߗׁj���}�X�^
        Public Const M_SpDays As String = "M_SpDays"            '�L�O���}�X�^
        Public Const M_KMealTp As String = "M_KMealTp"          '�����}�X�^�p�H��
        Public Const M_KMealLmt As String = "M_KMealLmt"        '�����}�X�^�p�H�퐧���l
        Public Const M_KMealLk As String = "M_KMealLk"          '�����}�X�^�p�H��֘A�}�X�^
        Public Const M_User As String = "M_User"                '���[�U�[�}�X�^
        Public Const M_UserGp As String = "M_UserGp"            '���[�U�[�O���[�v�}�X�^
        Public Const M_Color As String = "M_Color"              '�F���}�X�^
        Public Const M_ColorSet As String = "M_ColorSet"        '�F���ݒ�}�X�^
        Public Const T_MKonInf As String = "T_MKonInf"          '�����}�X�^
        Public Const T_MKonMk As String = "T_MKonMk"            '�����}�X�^�쐬
        Public Const T_MKonCk As String = "T_MKonCk"            '���������}�X�^
        Public Const T_MKonFd As String = "T_MKonFd"            '�����H�i�}�X�^
        Public Const T_KonMk As String = "T_KonMk"              '�ʏ팣���쐬��
        Public Const T_KonCk As String = "T_KonCk"              '�ʏ팣������
        Public Const T_KonFd As String = "T_KonFd"              '�ʏ팣���H�i
        Public Const T_CKonCk As String = "T_CKonCk"            '��֌�������
        Public Const T_CKonFd As String = "T_CKonFd"            '��֌����H�i
        Public Const T_CKonCnt As String = "T_CKonCnt"          '��֌����l��
        Public Const T_PerM As String = "T_PerM"                '�l�}�X�^
        Public Const T_PerHospi As String = "T_PerHospi"        '�l���މ@���
        Public Const T_PerInf As String = "T_PerInf"            '�l���
        Public Const T_PerEat As String = "T_PerEat"            '�l�H�����
        Public Const T_PerKCnt As String = "T_PerKCnt"          '�l���������
        Public Const T_PerMov As String = "T_PerMov"            '�l�ړ����
        Public Const T_PerDlv As String = "T_PerDlv"            '�l�z�V���
        Public Const T_PerCmt As String = "T_PerCmt"            '�l�R�����g
        Public Const T_PerLmt As String = "T_PerLmt"            '�l�����l
        Public Const T_PerMn As String = "T_PerMn"              '�l�������
        Public Const T_PerSel As String = "T_PerSel"            '�l�I��H
        Public Const T_PerNut As String = "T_PerNut"            '�l�o�ǉh�{
        Public Const T_PerMik As String = "T_PerMik"            '�l����
        Public Const T_PerVena As String = "T_PerVena"          '�l�Ö��h�{
        Public Const T_PerIntk As String = "T_PerIntk"          '�l�ێ�l
        Public Const T_PerTec As String = "T_PerTec"            '�l�h�{�w��
        Public Const T_PerIll As String = "T_PerIll"            '�l�a��
        Public Const T_MPCnt As String = "T_MPCnt"              '�\����{�l��
        Public Const T_HacInf As String = "T_HacInf"            '�������
        Public Const T_HacJZ As String = "T_HacJZ"              '�������݌ɏ��
        Public Const T_HacFd As String = "T_HacFd"              '�����H�i���
        Public Const T_HacHstNo As String = "T_HacHstNo"        '�������M�����}��
        Public Const T_HacHst As String = "T_HacHst"            '�������M����
        Public Const T_Buy As String = "T_Buy"                  '�d�����
        Public Const T_BuyInf As String = "T_BuyInf"            '�d���H�i���
        Public Const T_ZaiOut As String = "T_ZaiOut"            '�݌ɕ��o���
        Public Const T_Invty As String = "T_Invty"              '�I�����
        Public Const T_ZaiMony As String = "T_ZaiMony"          '�݌ɊȈՋ��z���
        Public Const T_ZaiSim As String = "T_ZaiSim"            '�݌ɊȈՏ��
        Public Const R_HVirusC As String = "R_HVirusC"          '���M�E�E�ۉ��H�L�^�뗿��
        Public Const R_HVirusF As String = "R_HVirusF"          '���M�E�E�ۉ��H�L�^��H�i
        Public Const R_HSeq As String = "R_HSeq"                '�������A�ԊǗ�
        Public Const R_EiyoS As String = "R_EiyoS"              '�h�{�o�[�N��
        Public Const W_CvtCd As String = "W_CvtCd"              '�ڍs�R�[�h�ϊ�WK
        Public Const W_5Nutri As String = "W_5Nutri"            '5���������[�N
        Public Const S_Frm As String = "S_Frm"                  '��ʃ}�X�^
        Public Const S_RptInf As String = "S_RptInf"            '���[�}�X�^
        Public Const S_ROutPut As String = "S_ROutPut"          '���[������}�X�^
        Public Const S_Menu As String = "S_Menu"                '���j���[�}�X�^
        Public Const S_Msg As String = "S_Msg"                  '���b�Z�[�W�}�X�^
        Public Const S_Sys As String = "S_Sys"                  '�V�X�e���ݒ�}�X�^
        Public Const S_SCK As String = "S_SCK"                  '�V���[�g�J�b�g�L�[�}�X�^
        Public Const S_Ctl As String = "S_Ctl"                  '�R���g���[���ݒ�}�X�^
        Public Const S_JigyoSnd As String = "S_JigyoSnd"        '���Ə��ݒ��ʃ}�X�^
    End Class

#Region "�����t�Ԓ�`"
    '**************************************************************************
    ' �����t�Ԓ�`
    ' ���t�ԏ����l�A�t�Ԑ����l�̓t�B�[���h�̌�����0���߂���1�ȏ�̐��l���L�q���ĉ������B
    ' ���t���̃e�[�u���́A���O�C�����̎{�݃R�[�h�������n���K�v������܂��B    
    '**************************************************************************
    Public Class AutoNumSec
        '�e�[�u������
        Public Const DBNAME_S_Menu As String = "S_Menu"             '���j���[�}�X�^
        Public Const DBNAME_M_Sisetsu As String = "M_Sisetsu"       '�{�݃}�X�^�e�[�u����(M_Sisetsu)
        Public Const DBNAME_M_Trader As String = "M_Trader"         '�d����}�X�^�e�[�u����(M_Trader)
        '----------------------------------------------------------------------------------------------�����U�ԑΉ�
        'Public Const DBNAME_M_Trader_H As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'H%')"         '�d����}�X�^�e�[�u����(�{�Зp)	2008/07/14
        'Public Const DBNAME_M_Trader_S As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'S%')"         '�d����}�X�^�e�[�u����(�x�X�p)	2008/07/14
        'Public Const DBNAME_M_Trader_J As String = "(select substring(TrndCd,2,7) as TrndCd from M_Trader where TrndCd like 'J%')"         '�d����}�X�^�e�[�u����(���Ə��p)	2008/07/14


        Public Const DBNAME_M_Fd As String = "M_Fd"                 '�H�i�}�X�^�e�[�u��
        'Public Const DBNAME_M_Fd_H As String = "(select substring(FdCd,2,7) as FdCd from M_Fd where substring(FdCd,1,1) = 'H')"       '�H�i�}�X�^�e�[�u��(�{�Зp)   (�ʏ�)     2009/09/11 �ǉ�
        'Public Const DBNAME_M_Fd_S As String = "(select substring(FdCd,2,7) as FdCd from M_Fd where substring(FdCd,1,1) = 'S')"       '�H�i�}�X�^�e�[�u��(�x�X�p)   (�ʏ�)     2009/09/11 �ǉ�
        Public Const DBNAME_M_MealTpK As String = "M_MealTpK"       '�H��}�X�^(����)�e�[�u��
        Public Const DBNAME_M_MealTpP As String = "M_MealTpP"       '�H��}�X�^(�l)�e�[�u��
        Public Const DBNAME_M_MealGpK As String = "M_MealGpK"       '�H��O���[�v�}�X�^(����)�e�[�u��
        Public Const DBNAME_M_MealGpP As String = "M_MealGpP"       '�H��O���[�v�}�X�^(�H��)�e�[�u��
        Public Const DBNAME_M_KMealTp As String = "M_KMealTp"       '�����}�X�^�p�H��e�[�u��
        'Public Const DBNAME_M_KMealTp_H As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'H')"       '�����}�X�^�p�H��e�[�u��(�{�Зp)
        'Public Const DBNAME_M_KMealTp_S As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'S')"       '�����}�X�^�p�H��e�[�u��(�x�X�p)
        'Public Const DBNAME_M_KMealTp_J As String = "(select substring(mealcd,2,4) as mealcd from M_KMealTp where substring(mealcd,1,1) = 'J')"       '�����}�X�^�p�H��e�[�u��(���Ə��p)
        Public Const DBNAME_M_Cok As String = "M_Cok"               '�����}�X�^�e�[�u��
        'Public Const DBNAME_M_Cok_H As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'H')"       '�����}�X�^�e�[�u��(�{�Зp)   (�ʏ�)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_S As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'S')"       '�����}�X�^�e�[�u��(�x�X�p)   (�ʏ�)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_J As String = "(select substring(CkCd,2,8) as CkCd from M_Cok where substring(CkCd,1,1) = 'J')"       '�����}�X�^�e�[�u��(���Ə��p) (�ʏ�)     2008/06/16 �ǉ�

        'Public Const DBNAME_M_Cok_HSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HSYU')" '�����}�X�^�e�[�u��(�{�Зp)   (��H)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_SSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SSYU')" '�����}�X�^�e�[�u��(�x�X�p)   (��H)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_JSYU As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JSYU')" '�����}�X�^�e�[�u��(���Ə��p) (��H)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_HKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HKKN')" '�����}�X�^�e�[�u��(�{�Зp)   (�o�ǉh�{) 2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_SKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SKKN')" '�����}�X�^�e�[�u��(�x�X�p)   (�o�ǉh�{) 2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_JKKN As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JKKN')" '�����}�X�^�e�[�u��(���Ə��p) (�o�ǉh�{) 2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_HMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'HMLK')" '�����}�X�^�e�[�u��(�{�Зp)   (����)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_SMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'SMLK')" '�����}�X�^�e�[�u��(�x�X�p)   (����)     2008/06/16 �ǉ�
        'Public Const DBNAME_M_Cok_JMLK As String = "(select substring(CkCd,5,5) as CkCd from M_Cok where substring(CkCd,1,4) = 'JMLK')" '�����}�X�^�e�[�u��(���Ə��p) (����)     2008/06/16 �ǉ�

        Public Const DBNAME_M_Cok_SYU As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'SYU')" '�����}�X�^�e�[�u��(��H)     2011/12/15 �ǉ�
        Public Const DBNAME_M_Cok_KKN As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'KKN')" '�����}�X�^�e�[�u��(�o�ǉh�{) 2011/12/15 �ǉ�
        Public Const DBNAME_M_Cok_MLK As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'MLK')" '�����}�X�^�e�[�u��(����)     2011/12/15 �ǉ�
        Public Const DBNAME_M_Cok_VEN As String = "(select substring(CkCd,6,6) as CkCd from M_Cok where substring(CkCd,1,3) = 'VEN')" '�����}�X�^�e�[�u��(�Ö��h�{)     2011/12/15 �ǉ�


        Public Const DBNAME_M_Age As String = "M_Age"               '�N��Ǘ��}�X�^�e�[�u��
        'Public Const DBNAME_M_Age_H As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'H')"       '�N��Ǘ��}�X�^�e�[�u��(�{�Зp)
        'Public Const DBNAME_M_Age_S As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'S')"       '�N��Ǘ��}�X�^�e�[�u��(�x�X�p)
        'Public Const DBNAME_M_Age_J As String = "(select substring(AgeCd,2,5) as AgeCd from M_Age where substring(AgeCd,1,1) = 'J')"       '�N��Ǘ��}�X�^�e�[�u��(���Ə��p)
        Public Const DBNAME_T_PerM As String = "T_PerM"             '�l�}�X�^
        Public Const DBNAME_M_Nutri As String = "M_Nutri"           '�����}�X�^
        'Public Const DBNAME_M_Nutri_H As String = "(select substring(NtCd,2,7) as NtCd,NtDiv from M_Nutri where substring(NtCd,1,1) = 'H')"   '�����}�X�^�e�[�u��(�{�Зp)   (�ʏ�)     2009/09/11 �ǉ�
        'Public Const DBNAME_M_Nutri_S As String = "(select substring(NtCd,2,7) as NtCd,NtDiv from M_Nutri where substring(NtCd,1,1) = 'S')"   '�����}�X�^�e�[�u��(�x�X�p)   (�ʏ�)     2009/09/11 �ǉ�
        Public Const DBNAME_A_Config As String = "A_Config"

        '�t�B�[���h��
        Public Const FIELDNAME_MENUNO As String = "MenuNo"              '�t�B�[���h��
        Public Const FIELDNAME_JICD As String = "JiCd"                  '�t�B�[���h���iJiCd�j
        Public Const FIELDNAME_TDCD As String = "TrndCd"                '�t�B�[���h���iTdCd�j
        Public Const FIELDNAME_FD_FDCD As String = "FdCd"               '�t�B�[���h��
        Public Const FIELDNAME_MEALTPK_MEALCD As String = "MealCd"      '�t�B�[���h��
        Public Const FIELDNAME_MEALTPP_MEALCD As String = "MealCd"      '�t�B�[���h��
        Public Const FIELDNAME_MMealGpK_MealGpCd As String = "MealGpCd" '�t�B�[���h���F�H��O���[�v�R�[�h
        Public Const FIELDNAME_MMEALGPP_MEALGPCD As String = "MealGpCd" '�t�B�[���h���F�H��O���[�v�R�[�h
        Public Const FIELDNAME_CKCD As String = "CkCd"                  '�t�B�[���h���iCkCd�j
        Public Const FIELDNAME_MAGE_AGECD As String = "AgeCd"           '�t�B�[���h���iAgeCd�j
        Public Const FIELDNAME_TPerM_PerCD As String = "PerCd"          '�t�B�[���h��(�l�R�[�h)
        Public Const FIELDNAME_MKMEALTP_MEALCD As String = "MealCd"     '�t�B�[���h��
        Public Const FIELDNAME_NtCd As String = "NtCd"                  '�t�B�[���h���i�����R�[�h�j

        '�t�ԏ����l
        Public Const INITVALUE_S_MENU As String = "0001"              '���j���[�}�X�^�t�ԏ����l
        Public Const INITVALUE_M_SISETSU As String = "0001"           '�{�݃}�X�^�t�ԏ����l

        Public Const INITVALUE_M_TRADER As String = "00000001"         '�d����}�X�^�t�ԏ����lAji7�Ή�
        'Public Const INITVALUE_M_TRADER As String = "0000001"         '�d����}�X�^�t�ԏ����l(�ύX�̉\������)
        Public Const LIMITVALUE_M_TRADER As String = "99999999"        '�d����}�X�^�t�Ԑ����lAji�V�Ή��W���ɕύX
        'Public Const LIMITVALUE_M_TRADER As String = "9999999"        '�d����}�X�^�t�Ԑ����l(�ύX�̉\������)����[HSJ]��t�^���邽�߂�7���ɕύX

        Public Const INITVALUE_M_FD As String = "00000001"            '�H�i�}�X�^�t�ԏ����l(�ύX�̉\������)
        Public Const LIMITVALUE_M_FD As String = "99999999"           '�H�i�}�X�^�t�Ԑ����l(�ύX�̉\������)

        Public Const INITVALUE_M_MEALTPK As String = "0001"           '�H��}�X�^(����)�t�ԏ����l(�ύX�̉\������)
        Public Const INITVALUE_M_MEALTPP As String = "0001"           '�H��}�X�^(�l)�t�ԏ����l(�ύX�̉\������)
        Public Const INITVALUE_M_MealGpK As String = "0001"           '�H��O���[�v�}�X�^(����)�t�ԏ����l(�ύX�̉\������)
        Public Const INITVALUE_M_MEALGPP As String = "0001"           '�H��O���[�v�}�X�^(�H��)�t�ԏ����l(�ύX�̉\������)

        Public Const INITVALUE_M_Cok As String = "000000001"           '�����}�X�^�t�ԏ����l
        'Public Const INITVALUE_M_Cok As String = "00000001"           '�����}�X�^�t�ԏ����l

        Public Const INITVALUE_M_CokNotNormal As String = "000001"     '�����}�X�^�t�ԏ����l(�ʏ�ȊO) 2011/12/15 �ǉ�
        'Public Const INITVALUE_M_CokNotNormal As String = "00001"     '�����}�X�^�t�ԏ����l(�ʏ�ȊO) 2008/06/16 �ǉ�

        Public Const INITVALUE_M_AGE As String = "00001"              '�N��Ǘ��}�X�^�t�ԏ����l
        Public Const INITVALUE_M_AGE_AUTH As String = "0001"          '�N��Ǘ��}�X�^�t�ԏ����l(����) 2008/06/17 �ǉ�
        Public Const INITVALUE_T_PerM As String = "0000000001"        '�l�}�X�^�t�ԏ����l(�ύX�̉\������)
        Public Const INITVALUE_M_KMEALTP As String = "0001"           '�����}�X�^�p�H��
        Public Const INITVALUE_M_KMEALTP_AUTH As String = "001"       '�����}�X�^�p�H��(����)
        Public Const INITVALUE_M_NtCd As String = "00000001"          '�����}�X�^�t�ԏ����l

        '�t�Ԑ����l
        Public Const LIMITVALUE_M_S_MENU As String = "9999"           '���j���[�}�X�^�t�Ԑ����l
        Public Const LIMITVALUE_M_SISETSU As String = "9999"          '�{�݃}�X�^�t�Ԑ����l



        Public Const LIMITVALUE_M_MEALTPK As String = "9999"          '�H��}�X�^(����)�t�Ԑ����l(�ύX�̉\������)
        Public Const LIMITVALUE_M_MEALTPP As String = "9999"          '�H��}�X�^(�l)�t�Ԑ����l(�ύX�̉\������)
        Public Const LIMITVALUE_M_MealGpK As String = "9999"          '�H��O���[�v�}�X�^(����)�t�Ԑ����l(�ύX�̉\������)
        Public Const LIMITVALUE_M_MEALGPP As String = "9999"          '�H��O���[�v�}�X�^(�H��)�t�Ԑ����l(�ύX�̉\������)


        Public Const LIMITVALUE_M_Cok As String = "999999999"          '�����}�X�^�t�Ԑ����l
        Public Const LIMITVALUE_M_CokNotNormal As String = "999999"    '�����}�X�^�t�Ԑ����l(�ʏ�ȊO) 2008/06/16 �ǉ�

        'Public Const LIMITVALUE_M_Cok As String = "99999999"          '�����}�X�^�t�Ԑ����l
        'Public Const LIMITVALUE_M_CokNotNormal As String = "99999"    '�����}�X�^�t�Ԑ����l(�ʏ�ȊO) 2008/06/16 �ǉ�

        Public Const LIMITVALUE_M_AGE As String = "99999"             '�N��Ǘ��}�X�^�t�Ԑ����l
        Public Const LIMITVALUE_M_AGE_AUTH As String = "9999"         '�N��Ǘ��}�X�^�t�Ԑ����l(����) 2008/06/17 �ǉ�
        Public Const LIMITVALUE_T_PerM As String = "9999999999"       '�l�}�X�^�t�Ԑ����l(�ύX�̉\������)

        Public Const LIMITVALUE_M_KMEALTP As String = "9999"          '�����}�X�^�p�H��
        Public Const LIMITVALUE_M_KMEALTP_AUTH As String = "999"      '�����}�X�^�p�H��(����)
        Public Const LIMITVALUE_M_NtCd As String = "99999999"         '�����}�X�^�t�Ԑ����l

    End Class

#End Region

#Region "���[�󎚊֘A"
    '**********************************************************************
    '���[�󎚊֘A
    '**********************************************************************
    Public Class Report

        Public Const PRINT_EXTENSION_TXT As String = ".txt"         'txt�t�@�C��
        Public Const PRINT_EXTENSION_CSV As String = ".csv"         'csv�t�@�C��
        Public Const PRINT_EXTENSION_RPT As String = ".rpt"         'rpt�t�@�C��
        Public Const PRINT_EXTENSION_PDF As String = ".pdf"         'pdf�t�@�C��
        Public Const PRINT_EXTENSION_XLS As String = ".xlsm"        'xls�t�@�C��
        Public Const PRINT_EXTENSION_ZIP As String = ".zip"         'zip�t�@�C��
        Public Const PRINT_EXTENSION_EXE As String = ".exe"         'exe�t�@�C��
        Public Const PRINT_EXTENSION_BAT As String = ".bat"         'bat�t�@�C��
        'Public Const PRINT_EXTENSION_A7R As String = ".a7r"         'a7r�t�@�C��
        Public Const PRINT_EXTENSION_A7U As String = ".a7u"         'a7u�t�@�C��(�A�b�v���[�h��)
        Public Const PRINT_EXTENSION_A7D As String = ".a7d"         'a7d�t�@�C��(�_�E�����[�h��)

        Public Const REPORT_FILE_HEADER_NAME As String = "Header"   'CSV�̃w�b�_�[�t�@�C����
        Public Const REPORT_FILE_DELIMITER As String = ">>>>> This line is Delimiter of ReportData for AjiDAS <<<<<"   '�d�����|�[�g�t�@�C���̃f���~�^
        Public Const REPORT_FILE_ERROR_INF As String = ">>>>> Error Information <<<<<"

        Public Const REPORT_FILE_MDB_NAME As String = "GTXPrint.mdb"
        Public Const REPORT_FILE_SCHEMA_NAME As String = "schema.ini"

    End Class
#End Region

#Region "�I�[�_�A�g�֘A"

    Public Class ORecv

        Public Const ORDER_DEFINE_ORDERTYPE_KAMEDA As String = "kameda"         '�T�c�I�[�_
        Public Const ORDER_DEFINE_ORDERTYPE_MIRAIS As String = "MIRAIs"         'MIRAIs�iNEC�ECSI�j
        Public Const ORDER_DEFINE_ORDERTYPE_FUJITSU_GX As String = "GX"         'GX�i�x�m�ʁj
        Public Const ORDER_DEFINE_ORDERTYPE_SOFTS As String = "SoftS"           '���(SoftS)
        Public Const ORDER_DEFINE_ORDERTYPE_PHRMAKER As String = "PHRMAKER"     'PHRMAKER(NJC)
        Public Const ORDER_DEFINE_ORDERTYPE_SBS As String = "SBS"               'SBS
        Public Const ORDER_DEFINE_ORDERTYPE_SOFT_STANDARD As String = "SoftD"   'SoftS�̕W��
        Public Const ORDER_DEFINE_ORDERTYPE_NAIS As String = "NAIS"             '�i�C�X
        Public Const ORDER_DEFINE_ORDERTYPE_ALICE As String = "ALICE"           '�A���X
        Public Const ORDER_DEFINE_ORDERTYPE_FUJITSU_EX_WEBEDITION As String = "EXWEB"   'EX WebEdition�i�x�m�ʁj
        Public Const ORDER_DEFINE_ORDERTYPE_MIC As String = "MIC"               'MIC ���f�B�J�����i���j
        Public Const ORDER_DEFINE_ORDERTYPE_SIEMENS As String = "SIEMS"               '�V�[�����X�@�V�[�����X�T�c��Ï��V�X�e��

        Public Const ORDER_DEFINE_UPDATE_USER As String = "ORDER"   '�I�[�_�A�g
        Public Const ORDER_DEFINE_UPDATE_KAMEDA_USER As String = "KAMEDA_ORDER"   '�I�[�_�A�g

        Public Const AUTO_EXEC_USER As String = "AUTOPG"            '����PG

        Public Const ORDER_TELMDIV_HEADER As String = "hd"           '�w�b�_�[����
        Public Const ORDER_TELMDIV_DATA As String = "data"           '�f�[�^����

        Public Const OrderRfctDiv_NotRfct As String = "0"           '������
        Public Const OrderRfctDiv_Rfct As String = "1"              '���f��
        Public Const OrderRfctDiv_RfctError As String = "9"         '���f�s��
        Public Const OrderRfctDiv_Cancel As String = "X"            '������
        Public Const OrderRfctDiv_Hold As String = "H"              '�ۗ�

        Public Const OrderChangeDivNm_ComExecDiv As String = "ComExecDiv"   '���ʏ����敪
        Public Const OrderChangeDivNm_ComOdDiv As String = "ComOdDiv"       '���ʍX�V�敪

        Public Enum enmOrderPpIsDiv
            NotPrt = 0          '�����s
            Prt = 1             '���s��
            NoPpIs = 9          '�ΏۊO
        End Enum

        Public Const OrderAtvOd_Ji As String = "1"              '���{�I�[�_
        Public Const OrderAtvOd_Oth As String = ""              '����ȊO

        Public Enum enmOrderExecDiv
            Insert = 1          '�V�K�ǉ�
            Update = 2          '�C��
            Delete = 3          '�폜
        End Enum

        Public Enum enmOrderNoFrm
            ShowFrm = 0             '��ʕ\��
            NotFrm = 1          '��ʔ�\��
        End Enum


        '���ʎw���敪
        Public Const OrderComOdDiv_InHospital As String = "INN"                 '���@
        Public Const OrderComOdDiv_OutHospital As String = "OUT"                '�މ@
        Public Const OrderComOdDiv_CancelInHospital As String = "DEI"           '���@���
        Public Const OrderComOdDiv_CancelOutHospital As String = "DEO"          '�މ@���
        Public Const OrderComOdDiv_ChangeInHospital As String = "INC"           '���@���ύX
        Public Const OrderComOdDiv_ChangeOutHospital As String = "OTC"          '�މ@���ύX
        Public Const OrderComOdDiv_CLINIC_DIALYSIS As String = "CLI"            '�O���E����

        Public Const OrderComOdDiv_UnLimitInHospital As String = "UIN"          '���������@�i���@���ł��G���[�Ƃ��Ȃ��j

        Public Const OrderComOdDiv_NoMeal As String = "NOM"                     '���H
        Public Const OrderComOdDiv_NoMealInHospital As String = "NIN"           '���H���@
        Public Const OrderComOdDiv_NoMealMove As String = "NMV"                 '���H�ړ�

        Public Const OrderComOdDiv_GoOut As String = "GOT"                      '�O�o
        Public Const OrderComOdDiv_SleepOut As String = "SLP"                   '�O��
        Public Const OrderComOdDiv_ComeBack As String = "CBK"                   '�A�@
        Public Const OrderComOdDiv_ComeBackEat As String = "CBE"                '�A�@+�H���ύX

        Public Const OrderComOdDiv_AllDataChange As String = "ALL"          '�S���ڕύX

        Public Const OrderComOdDiv_EatChange As String = "EAT"              '�H���ύX
        Public Const OrderComOdDiv_EatChangeRepeat As String = "EAR"        '�H���ύX(�A��)

        Public Const OrderComOdDiv_SelectEat As String = "SEL"              '�I��H

        Public Const OrderComOdDiv_Move As String = "MOV"                   '�ړ����X�V
        Public Const OrderComOdDiv_WardChange As String = "WRD"             '�]��
        Public Const OrderComOdDiv_RoomChange As String = "ROM"             '�]��
        Public Const OrderComOdDiv_MedicalChange As String = "MCA"          '�]��

        Public Const OrderComOdDiv_WardMedicalChange As String = "WMR"      '�]��+�]��
        Public Const OrderComOdDiv_RoomMedicalChange As String = "RMM"      '�]��+�]��

        Public Const OrderComOdDiv_DeliveryChange As String = "MDE"         '�z�V��ύX
        Public Const OrderComOdDiv_DeliveryAddExtensionMeal As String = "EML"      '���H


        Public Const OrderComOdDiv_PersonalInfo As String = "INF"           '�l���(�}�X�^�{INF�{ILL)
        Public Const OrderComOdDiv_ChangeDoctor As String = "DCT"           '�S���ҕϊ�

        Public Const OrderComOdDiv_TPerM As String = "PMS"                  '�l�}�X�^
        Public Const OrderComOdDiv_TPerEat As String = "PET"                '�l�H�����
        Public Const OrderComOdDiv_TPerMov As String = "PMV"                '�l�ړ����
        Public Const OrderComOdDiv_TPerInf As String = "PIF"                '�l���
        Public Const OrderComOdDiv_TPerDlv As String = "PDV"                '�l�z�V���
        Public Const OrderComOdDiv_TPerCmt As String = "PCM"                '�l�R�����g���
        Public Const OrderComOdDiv_TPerLmt As String = "PLT"                '�l�����l���
        Public Const OrderComOdDiv_TPerIll As String = "PIL"                '�l�a�����

        ''�x�m�ʃI�[�_�������
        Public Const OrderEGMAINGX_Ans_OK As String = "OK"                  'OK
        Public Const OrderEGMAINGX_Ans_NG As String = "NG"                  'NG�i�ڑ��ُ�j
        Public Const OrderEGMAINGX_Ans_N1 As String = "N1"                  'N1�i���g���C�j
        Public Const OrderEGMAINGX_Ans_N2 As String = "N2"                  'N2�i�X�L�b�v�j
        Public Const OrderEGMAINGX_Ans_N3 As String = "N3"                  'N3�i�_�E���j
        Public Const OrderEGMAINGX_Ans_N4 As String = "N4"                  'N4�i�f�[�^�Ȃ��j
        Public Const OrderEGMAINGX_Ans_N5 As String = "N5"                  'N5�i�f�[�^�Ȃ��j
        Public Const OrderEGMAINGX_Ans_N6 As String = "N6"                  'N6�i�X�V�G���[�j

        Public Const OrderEGMAINGX_TelMDiv_ME As String = "ME"                  'ME
        Public Const OrderEGMAINGX_TelMDiv_NI As String = "NI"                  'NI
        Public Const OrderEGMAINGX_TelMDiv_NS As String = "NS"                  'NS
        Public Const OrderEGMAINGX_TelMDiv_KJ As String = "KJ"                  'KJ
        Public Const OrderEGMAINGX_TelMDiv_CE As String = "CE"                  'CE(�h�{�A�g�p)

        '�ϊ�����
        Public Const OCHG_OrderMelCnvStr As String = "     "                    '�H��R�[�h��5�P�^
        Public Const OCHG_OrderMCkCnvStr As String = "     "                    '��H�R�[�h��5�P�^
        Public Const OCHG_OrderAgeCdCnvStr As String = "   "                    '�N��ʃR�[�h3�P�^
        Public Const OCHG_OrderWardCnvStr As String = "     "                   '�a���R�[�h5�P�^
        Public Const OCHG_OrderRoomCnvStr As String = "     "                   '�a���R�[�h5�P�^
        Public Const OCHG_OrderDelivCntStr As String = "     "                  '�z�V��R�[�h5�P�^
        Public Const OCHG_OrderCmtCnvStr As String = "        "                 '�R�����g�R�[�h8�P�^
        Public Const OCHG_OrderEatTmCnvInt As Integer = 0                       '���ԑс@�w��Ȃ�

        Public Const OCHG_OrderMealCnvSet As String = "MEAL"                    '�H��R�[�h
        Public Const OCHG_OrderMaCkCnvSet As String = "MACK"                    '��H�R�[�h
        Public Const OCHG_OrderAgeCnvSet As String = "AGE"                      '�N��R�[�h
        Public Const OCHG_OrderEatTmCnvSet As String = "EATTMDIV"               '���ԑ�
        Public Const OCHG_OrderCmt1CnvSet As String = "CMT1"                    '�R�����g�P
        Public Const OCHG_OrderCmt2CnvSet As String = "CMT2"                    '�R�����g�Q
        Public Const OCHG_OrderCmt3CnvSet As String = "CMT3"                    '�R�����g�R
        Public Const OCHG_OrderWardCnvSet As String = "WARD"                    '�a���R�[�h

        ''PHRMAKER�������
        Public Const OrderPHRMAKER_TelMDiv_PMS As String = "PMS"                  '���ҏ��
        Public Const OrderPHRMAKER_TelMDiv_INN As String = "UIN"                  '���������@���
        Public Const OrderPHRMAKER_TelMDiv_OUT As String = "OUT"                  '�މ@���
        Public Const OrderPHRMAKER_TelMDiv_EAT As String = "EAT"                  '�H�����
        Public Const OrderPHRMAKER_TelMDiv_MOV As String = "MOV"                  '�ړ����
        Public Const OrderPHRMAKER_TelMDiv_SCK As String = "SCK"                  '�a�����
        Public Const OrderPHRMAKER_TelMDiv_CMT As String = "CMT"                  '�R�����g���
        Public Const OrderPHRMAKER_TelMDiv_DAY As String = "DAY"                  '�ʉ@���
        Public Const OrderPHRMAKER_TelMDiv_NOM As String = "NOM"                  '��H���

        '�A�����M�[�敪
        Public Const OrderRsheetNm_Allergy As String = "ALLERGY"                '�A�����M�[����

        '�폜�̕����X�V�N����
        Public Const OrderDelYMDUpdDiv_NowDelYmd As String = "1"                '�폜�N�����ōX�V

        ''SBS�I�[�_���
        Public Const OrderSBS_TelMDiv_SL As String = "SL"                           'SL
        Public Const OrderSBS_TelMDiv_TB As String = "TB"                           'TB
        Public Const OrderSBS_TelMDiv_WH As String = "WH"                           'WH
        Public Const OrderSBS_TelMDiv_OD As String = "OD"                           'OD

        Public Const OrderSBS_TelMDiv_UPD As String = "UPD"                         'UPD
        Public Const OrderSBS_TelMDiv_UPWH As String = "UPWH"                       'UPWH

        ''SOFTS�I�[�_���
        Public Const OrderSOFT_TelMDiv_CSV As String = "CSV"                        'CSV

        ''�i�C�X�I�[�_���
        Public Const OrderNAIS_TelMDiv_P As String = "P"    '���ҏ�񃌃R�[�h
        Public Const OrderNAIS_TelMDiv_K As String = "K"    '���Ғ��Ӄ��R�[�h
        Public Const OrderNAIS_TelMDiv_T As String = "T"    '���H���ԃ��R�[�h
        Public Const OrderNAIS_TelMDiv_D As String = "D"    '���H�ڍ׃��R�[�h

        ''�A���X�I�[�_���
        Public Const OrderAlice_TelMDiv_TEXT As String = "TEXT"                        'TEXT

        ''MIC�I�[�_���
        Public Const OrderMIC_TelMDiv_TEXT As String = "TEXT"                        'TEXT

        ''SIEMENS�I�[�_���
        Public Const OrderSIEMS_TelMDiv_HEAD As String = "HEAD"                        '�w�b�_�[����
        Public Const OrderSIEMS_TelMDiv_TEXT As String = "TEXT"                        '�w�b�_�[�������e�L�X�g����

    End Class

#End Region

    '**********************************************************************
    '���̑�
    '**********************************************************************
    Public Const DelimiterChar As Char = ":"c          '�R���{���X�g���ł̃R�[�h�����̂̋�؂蕶��
    Public Const REG_AJIDAS7 As String = "SOFTWARE\FUJITELECOM\AjiDAS7"     'AjiDAS7�̃��W�X�g���L�[
    Public Const REG_NAME_L As String = "L"     '���C�Z���X�F�؂̃��W�X�g����
End Class
