Attribute VB_Name = "common"
Option Explicit
Public Declare Function InternetGetConnectedState Lib "wininet.dll" _
    (ByRef dwflags As Long, ByVal dwReserved As Long) As Long

'lpdwFlags
Public Const INTERNET_CONNECTION_MODEM As Long = 1         '接続にモデムを使用
Public Const INTERNET_CONNECTION_LAN As Long = 2           '接続にLANを使用
Public Const INTERNET_CONNECTION_PROXY As Long = 4         '接続にプロキシ・サーバーを使用
Public Const INTERNET_CONNECTION_MODEM_BUSY As Long = 8    '何も使用されていない
Public Const INTERNET_RAS_INSTALLED As Long = 16           'RASがインストールされている
Public Const INTERNET_CONNECTION_OFFLINE As Long = 32      'オフライン
Public Const INTERNET_CONNECTION_CONFIGURED As Long = 64   '有効な接続があるが現在接続されていない

Public Const GC_APLI_NAME = "コンピ取得ツール"
Public Const GC_THANKS = "をご利用いただき、ありがとうございます！"
Public Const GC_AMAZON = "<a href=""http://www.amazon.co.jp/?_encoding=UTF8&camp=247&creative=1211&linkCode=as2&tag=derb-22"">Amazonでのお買い物は、こちらから！開発支援にご協力お願いします。オーケー馬（半製品版）も絶賛、発売中！よろしくね</a>"
Public Const GC_BLOG_MAIL = "a585c4de0e448f@mo.jugem.jp"
Public Const GC_MAC_MAIL = "racesoft@buhi-buhi.com"
Public Const GC_FAIL_MAIL = "利用確認メールが送信できませんでした。"

'API
'INIファイル関連
'=============================================================
' INIファイルから指定したセクション名、キー名の値(数値)を取得
'=============================================================
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'===============================================================
' INIファイルから指定したセクション名、キー名へ値(文字列)を格納
'===============================================================
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public mode As Integer


'定数
'戻り値
Public Const G_OK = 1
Public Const G_NG = 0

'拡張子
Public Const G_EXTEND_INI = ".ini"

'INIファイル関連
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

'ファイル関連
Public Const G_COMMON_INIFILE = "keiba"       '共通INIファイル

'変数
Public gRet As Integer                      '戻り値
Public gbRet As Boolean                      '戻り値
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


'グリッド(芝、ダート)
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


'関数
'INIファイル関連
'====================================
' INIファイルへ値(文字列)を格納する
'=====================================
Public Function saveIni(pFile As String, pSec As String, pKey As String, pValue As String) As Integer
    Dim result As Long '戻り値（0=失敗、0<>成功)
    Dim filename As String
    
    saveIni = G_NG
    

    filename = App.Path & "\" & pFile & G_EXTEND_INI
    
    result = WritePrivateProfileString(pSec, pKey, pValue, filename)

    saveIni = result
End Function

'====================================
' INIファイルから値(文字列)を取得する
'=====================================
Public Function loadIni(pFile As String, pSec As String, pKey As String, pValue As String) As Integer
    Dim lpReturnedString As String * 1024 '格納バッファ
    Dim result As Long    '戻り値 (取得した値の文字数)
    Dim filename As String
    
    loadIni = G_NG
    
    filename = App.Path & "\" & pFile & G_EXTEND_INI
    result = GetPrivateProfileString(pSec, pKey, "", lpReturnedString, Len(lpReturnedString), filename)
    pValue = Left(lpReturnedString, InStr(lpReturnedString, Chr(0)) - 1)
    
    loadIni = result
End Function

'----- BubbleSort -----
'指定された配列内の要素を、隣接交換法(バブルソート)によりソートします。
'
'引数 lngArray()
'   ソートを行いたい配列を指定します。例えば要素が
'       lngArray(0) = 8
'       lngArray(1) = 2
'       lngArray(2) = 5
'   の配列を渡した場合、
'       lngArray(0) = 2
'       lngArray(1) = 5
'       lngArray(2) = 8
'   のように正順に整列されます。
'
'引数 lngStart
'   省略可能です。ソートを開始したい要素の番号を指定します。
'   省略した場合は引数 lngArray() の最小要素番号からソートを行います。
'
'引数 lngEnd
'   省略可能です。ソートを終了したい要素の番号を指定します。
'   省略した場合は引数 lngArray() の最大要素番号までソートを行います。
'
Public Sub BubbleSort _
    (ByRef lngArray() As Long, _
     Optional ByVal lngStart As Long, _
     Optional ByVal lngEnd As Long)

 Dim i As Long                                              'ループカウンタ
 Dim j As Long                                              'ループカウンタ
 Dim w As Long                                              '作業域
 
    If Not CBool(lngStart) Then lngStart = LBound(lngArray) '開始要素番号が指定されていない場合、最小要素番号を求める
    If Not CBool(lngEnd) Then lngEnd = UBound(lngArray)     '終了要素番号が指定されていない場合、最大要素番号を求める
    If lngStart >= lngEnd Then Exit Sub                     '終了要素番号が開始要素番号以下の場合、プロシージャを抜ける
    
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
'    Call GetTanPutFile(1, "c:\test\" & Format$(Now, "yyyymmddmmnnss") & ".mailtxt", Format$(Now, "yyyymmddHHnnss") & " 予想作業中でございます・・・")
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
'    'TargetUmaと合致する馬が１０番人気以下（１１番など）なら、購入対象(grdTargetに表示)
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
        jyoName = "札幌"
    Case "02"
        jyoName = "函館"
    Case "03"
        jyoName = "福島"
    Case "04"
        jyoName = "新潟"
    Case "05"
        jyoName = "東京"
    Case "06"
        jyoName = "中山"
    Case "07"
        jyoName = "中京"
    Case "08"
        jyoName = "京都"
    Case "09"
        jyoName = "阪神"
    Case "10"
        jyoName = "小倉"
    End Select
    
    gCnvJyoCD2JyoName = jyoName
    
End Function


''LAN接続状態確認
'Sub LAN接続_Status()
'    Dim lngState    As Long
'    Dim Ret         As String
'    '
'    Ret = API_InternetGetConnectedState()
'    MsgBox Ret
'End Sub

'LAN接続状態を求める
Function API_InternetGetConnectedState() As Integer
    Dim lngFlg      As Long
    Dim blnRet      As Boolean
    Dim strMsg      As String
    Dim iRet As Integer
    
    iRet = 1
    
    '
    On Error GoTo ErrorHandler
    '
    'InternetGetConnectedState API 関数呼出
    blnRet = InternetGetConnectedState(lngFlg, 0)
    '
    '接続結果の取得（ビット比較）
    If blnRet Then
        '(接続時)
        If lngFlg And INTERNET_CONNECTION_CONFIGURED Then
'            strMsg = strMsg & "有効な接続がありますが、現在接続されていません。" & vbCrLf
        End If
        If lngFlg And INTERNET_CONNECTION_LAN Then
'            strMsg = strMsg & "LAN経由でインターネットに接続されています。" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_PROXY Then
'            strMsg = strMsg & "プロキシ・サーバーを使用しています。" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_MODEM Then
'            strMsg = strMsg & "モデムを使用してインターネットに接続しています。" & vbCrLf
            iRet = 0
        End If
        If lngFlg And INTERNET_CONNECTION_OFFLINE Then
'            strMsg = strMsg & "現在オフラインです。" & vbCrLf
        End If
        If lngFlg And INTERNET_CONNECTION_MODEM_BUSY Then
'            strMsg = strMsg & "現在インターネット接続以外にモデムが使用されています。" & vbCrLf
        End If
        If lngFlg And INTERNET_RAS_INSTALLED Then
'            strMsg = strMsg & "リモート・アクセス・サーバーがインストールされています。" & vbCrLf
        End If
    Else
        '(オフライン時)
'        strMsg = "現在インターネットには接続していません。" & vbCrLf
    End If
    '
    '戻り値
'    API_InternetGetConnectedState = Left(strMsg, Len(strMsg) - 2)
    API_InternetGetConnectedState = iRet
    Exit Function

ErrorHandler:
    MsgBox Err.Number & vbTab & Err.Description & vbTab, vbOKOnly + vbExclamation, "InternetGetConnectedState API 関数"
           
End Function


