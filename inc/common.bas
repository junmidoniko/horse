Attribute VB_Name = "common"
Option Explicit
Public Declare Function InternetGetConnectedState Lib "wininet.dll" _
    (ByRef dwflags As Long, ByVal dwReserved As Long) As Long

'lpdwFlags
Public Const INTERNET_CONNECTION_MODEM As Long = 1         '�ڑ��Ƀ��f�����g�p
Public Const INTERNET_CONNECTION_LAN As Long = 2           '�ڑ���LAN���g�p
Public Const INTERNET_CONNECTION_PROXY As Long = 4         '�ڑ��Ƀv���L�V�E�T�[�o�[���g�p
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = 8    '�����g�p����Ă��Ȃ�
Public Const INTERNET_RAS_INSTALLED As Long = 16           'RAS���C���X�g�[������Ă���
Public Const INTERNET_CONNECTION_OFFLINE As Long = 32      '�I�t���C��
Public Const INTERNET_CONNECTION_CONFIGURED As Long = 64   '�L���Ȑڑ������邪���ݐڑ�����Ă��Ȃ�

Public Const GC_APLI_NAME = "�R���s�擾�c�[��"
Public Const GC_THANKS = "�������p���������A���肪�Ƃ��������܂��I"
Public Const GC_AMAZON = "<a href=""http://www.amazon.co.jp/?_encoding=UTF8&camp=247&creative=1211&linkCode=as2&tag=derb-22"">Amazon�ł̂��������́A�����炩��I�J���x���ɂ����͂��肢���܂��B�I�[�P�[�n�i�����i�Łj����^�A�������I��낵����</a>"
Public Const GC_BLOG_MAIL = "a585c4de0e448f@mo.jugem.jp"
Public Const GC_MAC_MAIL = "racesoft@buhi-buhi.com"
Public Const GC_FAIL_MAIL = "���p�m�F���[�������M�ł��܂���ł����B"

'API
'INI�t�@�C���֘A
'=============================================================
' INI�t�@�C������w�肵���Z�N�V�������A�L�[���̒l(���l)���擾
'=============================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'===============================================================
' INI�t�@�C������w�肵���Z�N�V�������A�L�[���֒l(������)���i�[
'===============================================================
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public mode As Integer


'�萔
'�߂�l
Public Const G_OK = 1
Public Const G_NG = 0

'�g���q
Public Const G_EXTEND_INI = ".ini"

'INI�t�@�C���֘A
Public Const G_INI_SEC_WINDOW = "Window"
Public Const G_INI_KEY_WINDOW_TOP = "TOP"
Public Const G_INI_KEY_WINDOW_LEFT = "LEFT"
Public Const G_INI_KEY_WINDOW_ALL_TOP = "TOP_all"
Public Const G_INI_KEY_WINDOW_ALL_LEFT = "LEFT_all"
Public Const G_INI_KEY_SIZE_ALL_HEIGHT = "height_all"
Public Const G_INI_KEY_SIZE_ALL_WIDTH = "width_all"
Public Const G_INI_SEC_CMPI = "Cmpi"
Public Const G_INI_KEY_CMPI_TXT = "TXT"
Public Const G_INI_SEC_DB = "DB"
Public Const G_INI_KEY_DB_FROMTIME = "fromTime"

'�t�@�C���֘A
Public Const G_COMMON_INIFILE = "keiba"       '����INI�t�@�C��

'�ϐ�
Public gRet As Integer                      '�߂�l
Public gbRet As Boolean                      '�߂�l
Public gHappyoTime As String

Public g_col As Long
Public Const COL1 = &HFFFFFF
Public Const COL2 = &HC0C0C0
Public Const COL3 = &H808080
Public Const COL4 = &H1
Public Const COL5 = &H800000
Public Const COL6 = &HFF0000
Public Const COL7 = &H800080
Public Const COL8 = &HFF00FF
Public Const COL9 = &H18000
Public Const COL10 = &H1FF00
Public Const COL11 = &H808000
Public Const COL12 = &HFFFF00
Public Const COL13 = &H80
Public Const COL14 = &HFF
Public Const COL15 = &H18080
Public Const COL16 = &H1FFFF


'�O���b�h(�ŁA�_�[�g)
Public Const GRD_SHIBA = 0
Public Const GRD_DART = 1

Public Enum grd_detail
    dyear = 0
    dmonthday
    djyocd
    dracenum
    dumaban
    dkettonum
    dbamei
    dracenum2
    dchokyosicode
    dkisyucode
    dkisyuryakusyo
    dkakuteijyuni
    dkisyurank
    dkisyuvalue
    dtanninki
    dfuku
    dharonrank
    dspace_jk
    dspace_odds
    dspace_haron
    dspace_jiseki
    dspace_free
    ddatas
End Enum


'�֐�
'INI�t�@�C���֘A
'====================================
' INI�t�@�C���֒l(������)���i�[����
'=====================================
Public Function saveIni(pFile As String, pSec As String, pKey As String, pValue As String) As Integer
    Dim result As Long '�߂�l�i0=���s�A0<>����)
    Dim filename As String
    
    saveIni = G_NG
    

    filename = App.Path & "\" & pFile & G_EXTEND_INI
    
    result = WritePrivateProfileString(pSec, pKey, pValue, filename)

    saveIni = result
End Function

'====================================
' INI�t�@�C������l(������)���擾����
'=====================================
Public Function loadIni(pFile As String, pSec As String, pKey As String, pValue As String) As Integer
    Dim lpReturnedString As String * 1024 '�i�[�o�b�t�@
    Dim result As Long    '�߂�l (�擾�����l�̕�����)
    Dim filename As String
    
    loadIni = G_NG
    
    filename = App.Path & "\" & pFile & G_EXTEND_INI
    result = GetPrivateProfileString(pSec, pKey, "", lpReturnedString, Len(lpReturnedString), filename)
    pValue = Left(lpReturnedString, InStr(lpReturnedString, Chr(0)) - 1)
    
    loadIni = result
End Function

'----- BubbleSort -----
'�w�肳�ꂽ�z����̗v�f���A�אڌ����@(�o�u���\�[�g)�ɂ��\�[�g���܂��B
'
'���� lngArray()
'   �\�[�g���s�������z����w�肵�܂��B�Ⴆ�Ηv�f��
'       lngArray(0) = 8
'       lngArray(1) = 2
'       lngArray(2) = 5
'   �̔z���n�����ꍇ�A
'       lngArray(0) = 2
'       lngArray(1) = 5
'       lngArray(2) = 8
'   �̂悤�ɐ����ɐ��񂳂�܂��B
'
'���� lngStart
'   �ȗ��\�ł��B�\�[�g���J�n�������v�f�̔ԍ����w�肵�܂��B
'   �ȗ������ꍇ�͈��� lngArray() �̍ŏ��v�f�ԍ�����\�[�g���s���܂��B
'
'���� lngEnd
'   �ȗ��\�ł��B�\�[�g���I���������v�f�̔ԍ����w�肵�܂��B
'   �ȗ������ꍇ�͈��� lngArray() �̍ő�v�f�ԍ��܂Ń\�[�g���s���܂��B
'
Public Sub BubbleSort _
    (ByRef lngArray() As Long, _
     Optional ByVal lngStart As Long, _
     Optional ByVal lngEnd As Long)

 Dim i As Long                                              '���[�v�J�E���^
 Dim j As Long                                              '���[�v�J�E���^
 Dim w As Long                                              '��ƈ�
 
    If Not CBool(lngStart) Then lngStart = LBound(lngArray) '�J�n�v�f�ԍ����w�肳��Ă��Ȃ��ꍇ�A�ŏ��v�f�ԍ������߂�
    If Not CBool(lngEnd) Then lngEnd = UBound(lngArray)     '�I���v�f�ԍ����w�肳��Ă��Ȃ��ꍇ�A�ő�v�f�ԍ������߂�
    If lngStart >= lngEnd Then Exit Sub                     '�I���v�f�ԍ����J�n�v�f�ԍ��ȉ��̏ꍇ�A�v���V�[�W���𔲂���
    
    i = lngEnd
    Do Until i <= lngStart
        j = lngStart
        Do Until j >= i
            If lngArray(j) >= lngArray(j + 1) Then
                w = lngArray(j)
                lngArray(j) = lngArray(j + 1)
                lngArray(j + 1) = w
            End If
            j = j + 1
        Loop
        i = i - 1
    Loop

End Sub
'pIdx 0:year, 1:monthday
Public Function getDate(pDat As String, pIdx As Integer) As String
    Select Case pIdx
    Case 0
        getDate = Left$(pDat, 4)
    Case 1
        getDate = Right$(pDat, 4)
    End Select
End Function
Public Sub soundBeep()
'Call BeepAPI(262, 300)
'Call BeepAPI(294, 300)
'Call BeepAPI(330, 300)
'Call BeepAPI(349, 300)
'Call BeepAPI(392, 300)
'Call BeepAPI(440, 300)
'Call BeepAPI(494, 300)
'Call BeepAPI(523, 300)
'    Call GetTanPutFile(1, "c:\test\" & Format$(Now, "yyyymmddmmnnss") & ".mailtxt", Format$(Now, "yyyymmddHHnnss") & " �\�z��ƒ��ł������܂��E�E�E")
End Sub


Public Function GetTanPutFile(pMode As Integer, pFile As String, pData As String) As Boolean
    Dim cnt As Integer
    Dim i As Integer
    Dim fn As Integer
    Dim wk As String
    Dim retry_cnt As Integer
    
    On Error GoTo err_handler
    
    GetTanPutFile = False
    
'    retry_cnt = 0
'
'    gstrSql = ""
'    gstrSql = gstrSql + "select Umaban, TanNinki "
'    gstrSql = gstrSql + "from odds_tanpuku_JK Where "
'    gstrSql = gstrSql + "Year = '" & gYear & "' and "
'    gstrSql = gstrSql + "MonthDay = '" & gMonthDay & "' and "
'    gstrSql = gstrSql + "JyoCD = '" & gJyoCD & "' and "
'    gstrSql = gstrSql + "RaceNum = '" & gRaceNum & "' "
'    gstrSql = gstrSql + "order by Umaban"
'retry_1:
'    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
'
'    cnt = -1
'    Do
'        If Rs.EOF = False Then
'            cnt = cnt + 1
'            ReDim Preserve gTanNinki(cnt)
'            ReDim Preserve gTanUmaban(cnt)
'
'            gTanUmaban(cnt) = Rs("Umaban")
'            gTanNinki(cnt) = Rs("TanNinki")
'
'            Rs.MoveNext
'
'            GetTanPutFile = True
'        Else
'            Exit Do
'        End If
'    Loop
'
'    Rs.Close
'
'    If cnt = -1 Then
'        retry_cnt = retry_cnt + 1
'        If retry_cnt < 10 Then
'            Sleep (1000)
'            Debug.Print "retry=" & retry_cnt
'            GoTo retry_1
'        Else
'            Exit Function
'        End If
'    End If
'
'    'TargetUma�ƍ��v����n���P�O�Ԑl�C�ȉ��i�P�P�ԂȂǁj�Ȃ�A�w���Ώ�(grdTarget�ɕ\��)
'    Call dspTarget(pMode)
'
'    If GetTanPutFile = True And pMode = 1 Then
'        fn = FreeFile
'        Open "c:\" & gYear & gMonthDay & gJyoCD & gRaceNum & ".txt" For Output As #fn
'
'        For i = 0 To cnt
'            wk = gTanUmaban(i) & "," & gTanNinki(i)
'            Print #fn, wk
'        Next i
'
'        Close #fn
'    End If
    
    fn = FreeFile
    Open pFile For Output As #fn
    Print #fn, pData
    Close #fn
    GetTanPutFile = True
    
    Exit Function

err_handler:
    MsgBox Err.Description & ":GetTanPutFile"
    Exit Function
End Function

Public Function gCnvJyoCD2JyoName(JyoCD As String) As String
    Dim jyoName  As String
    
    Select Case JyoCD
    Case "01"
        jyoName = "�D�y"
    Case "02"
        jyoName = "����"
    Case "03"
        jyoName = "����"
    Case "04"
        jyoName = "�V��"
    Case "05"
        jyoName = "����"
    Case "06"
        jyoName = "���R"
    Case "07"
        jyoName = "����"
    Case "08"
        jyoName = "���s"
    Case "09"
        jyoName = "��_"
    Case "10"
        jyoName = "���q"
    End Select
    
    gCnvJyoCD2JyoName = jyoName
    
End Function


''LAN�ڑ���Ԋm�F
'Sub LAN�ڑ�_Status()
'    Dim lngState    As Long
'    Dim Ret         As String
'    '
'    Ret = API_InternetGetConnectedState()
'    MsgBox Ret
'End Sub

'LAN�ڑ���Ԃ����߂�
Function API_InternetGetConnectedState() As Integer
    Dim lngFlg      As Long
    Dim blnRet      As Boolean
    Dim strMsg      As String
    Dim iRet As Integer
    
    iRet = 1
    
    '
    On Error GoTo ErrorHandler
    '
    'InternetGetConnectedState API �֐��ďo
    blnRet = InternetGetConnectedState(lngFlg, 0)
    '
    '�ڑ����ʂ̎擾�i�r�b�g��r�j
    If blnRet Then
        '(�ڑ���)
        If lngFlg And INTERNET_CONNECTION_CONFIGURED Then
'            strMsg = strMsg & "�L���Ȑڑ�������܂����A���ݐڑ�����Ă��܂���B" & vbCrLf
        End If
        If lngFlg And INTERNET_CONNECTION_LAN Then
'            strMsg = strMsg & "LAN�o�R�ŃC���^�[�l�b�g�ɐڑ�����Ă��܂��B" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_PROXY Then
'            strMsg = strMsg & "�v���L�V�E�T�[�o�[���g�p���Ă��܂��B" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_MODEM Then
'            strMsg = strMsg & "���f�����g�p���ăC���^�[�l�b�g�ɐڑ����Ă��܂��B" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_OFFLINE Then
'            strMsg = strMsg & "���݃I�t���C���ł��B" & vbCrLf
        End If
        If lngFlg And INTERNET_CONNECTION_MODEM_BUSY Then
'            strMsg = strMsg & "���݃C���^�[�l�b�g�ڑ��ȊO�Ƀ��f�����g�p����Ă��܂��B" & vbCrLf
        End If
        If lngFlg And INTERNET_RAS_INSTALLED Then
'            strMsg = strMsg & "�����[�g�E�A�N�Z�X�E�T�[�o�[���C���X�g�[������Ă��܂��B" & vbCrLf
        End If
    Else
        '(�I�t���C����)
'        strMsg = "���݃C���^�[�l�b�g�ɂ͐ڑ����Ă��܂���B" & vbCrLf
    End If
    '
    '�߂�l
'    API_InternetGetConnectedState = Left(strMsg, Len(strMsg) - 2)
    API_InternetGetConnectedState = iRet
    Exit Function

ErrorHandler:
    MsgBox Err.Number & vbTab & Err.Description & vbTab, vbOKOnly + vbExclamation, "InternetGetConnectedState API �֐�"
           
End Function


