VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "compi 20180812"
   ClientHeight    =   1455
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   2820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2820
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton Command48 
      Caption         =   "Command48"
      Height          =   495
      Left            =   11280
      TabIndex        =   67
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command47 
      Caption         =   "Command47"
      Height          =   615
      Left            =   11160
      TabIndex        =   66
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Command46"
      Height          =   615
      Left            =   11400
      TabIndex        =   65
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Command44"
      Height          =   375
      Left            =   9600
      TabIndex        =   64
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command44 
      Caption         =   "3連複"
      Height          =   375
      Left            =   9600
      TabIndex        =   63
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   10800
      TabIndex        =   62
      Text            =   "3"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3240
      Top             =   4440
   End
   Begin VB.CommandButton Command43 
      Caption         =   "巡回"
      Height          =   375
      Left            =   9480
      TabIndex        =   61
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   9480
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   60
      Text            =   "Form1.frx":0000
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton Command42 
      Caption         =   "yahoo odds hasso"
      Height          =   255
      Left            =   9480
      TabIndex        =   59
      Top             =   2280
      Width           =   2415
   End
   Begin VB.CommandButton Command41 
      Caption         =   "List2Site"
      Height          =   525
      Left            =   5760
      TabIndex        =   58
      Top             =   7560
      Width           =   1545
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Site2Data"
      Height          =   525
      Left            =   5760
      TabIndex        =   57
      Top             =   8280
      Width           =   1545
   End
   Begin VB.CommandButton Command39 
      Caption         =   "URL List"
      Height          =   645
      Left            =   5760
      TabIndex        =   56
      Top             =   6840
      Width           =   1545
   End
   Begin VB.ListBox List1 
      Height          =   960
      Left            =   3480
      OLEDropMode     =   1  '手動
      TabIndex        =   55
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command38 
      Caption         =   "cmpiDBget"
      Height          =   405
      Left            =   2760
      TabIndex        =   54
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command37 
      Caption         =   "patch1"
      Height          =   495
      Left            =   5880
      TabIndex        =   53
      Top             =   8880
      Width           =   1215
   End
   Begin VB.CommandButton Command36 
      Caption         =   "check2"
      Height          =   495
      Left            =   8160
      TabIndex        =   52
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "check"
      Height          =   495
      Left            =   8160
      TabIndex        =   51
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "yahooResGet"
      Height          =   495
      Left            =   9600
      TabIndex        =   50
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      Caption         =   "res_Match"
      Height          =   495
      Left            =   9600
      TabIndex        =   49
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "res_List"
      Height          =   495
      Left            =   9600
      TabIndex        =   48
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "yahoo_res"
      Height          =   495
      Left            =   8280
      TabIndex        =   47
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   ".netkeiba_home"
      Height          =   495
      Left            =   9600
      TabIndex        =   46
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "cmpiIn"
      Height          =   495
      Left            =   7920
      TabIndex        =   45
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "haraiALL"
      Height          =   495
      Left            =   7920
      TabIndex        =   44
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "hrai"
      Height          =   495
      Left            =   7920
      TabIndex        =   43
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command26"
      Height          =   495
      Left            =   8280
      TabIndex        =   42
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   5040
      TabIndex        =   41
      Text            =   "C:\2017\src\siteDL\"
      Top             =   480
      Width           =   6255
   End
   Begin VB.CommandButton Command25 
      Caption         =   "release"
      Height          =   495
      Left            =   1680
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   37
      Top             =   2880
      Width           =   1455
      Begin VB.OptionButton optMode 
         Caption         =   "コンピ指数"
         Height          =   255
         Index           =   1
         Left            =   240
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   39
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optMode 
         Caption         =   "払い戻し"
         Height          =   255
         Index           =   0
         Left            =   240
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Command24"
      Height          =   495
      Left            =   4680
      TabIndex        =   36
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H0000C0C0&
      Caption         =   "rakuten all harai"
      Height          =   645
      Left            =   5760
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   35
      Top             =   6210
      Width           =   1545
   End
   Begin VB.CommandButton Command22 
      Caption         =   "nankan result"
      Height          =   645
      Left            =   5730
      TabIndex        =   34
      Top             =   5280
      Width           =   1545
   End
   Begin VB.TextBox txtUrl 
      Height          =   435
      Left            =   5760
      TabIndex        =   33
      Text            =   "http://race.netkeiba.com/?pid=yoso_cp&id=c201710010201"
      Top             =   4380
      Width           =   5715
   End
   Begin VB.CommandButton Command21 
      Caption         =   "netKeiba.com2"
      Height          =   645
      Left            =   5760
      TabIndex        =   32
      Top             =   3690
      Width           =   1545
   End
   Begin VB.CommandButton Command20 
      Caption         =   "netKeiba.com"
      Height          =   645
      Left            =   5820
      TabIndex        =   31
      Top             =   2790
      Width           =   1545
   End
   Begin VB.CommandButton Command19 
      Caption         =   "test"
      Height          =   645
      Left            =   5820
      TabIndex        =   30
      Top             =   1920
      Width           =   1545
   End
   Begin VB.CommandButton Command18 
      Caption         =   "rakuten odds"
      Height          =   645
      Left            =   5850
      TabIndex        =   29
      Top             =   870
      Width           =   1545
   End
   Begin VB.CommandButton Command17 
      Caption         =   "today"
      Height          =   315
      Left            =   1620
      TabIndex        =   28
      Top             =   2580
      Width           =   825
   End
   Begin VB.CommandButton Command16 
      Caption         =   "cmpiList"
      Height          =   405
      Left            =   2820
      TabIndex        =   27
      Top             =   2070
      Width           =   1455
   End
   Begin VB.CommandButton Command15 
      Caption         =   "clear"
      Height          =   315
      Left            =   1620
      TabIndex        =   26
      Top             =   2220
      Width           =   855
   End
   Begin VB.CheckBox chkDL 
      Caption         =   "DL"
      Height          =   525
      Left            =   30
      TabIndex        =   25
      Top             =   2220
      Width           =   1245
   End
   Begin VB.CheckBox chkD 
      Caption         =   "Dsp"
      Height          =   525
      Left            =   3600
      TabIndex        =   24
      Top             =   3000
      Value           =   1  'ﾁｪｯｸ
      Width           =   1245
   End
   Begin VB.CommandButton Command14 
      Caption         =   "中央コンピデータ取得_res"
      Height          =   525
      Left            =   2790
      TabIndex        =   23
      Top             =   1470
      Width           =   1485
   End
   Begin VB.TextBox areaMD 
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   22
      Top             =   1080
      Width           =   675
   End
   Begin VB.TextBox areaMD 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   20
      Top             =   720
      Width           =   675
   End
   Begin VB.TextBox areaY 
      Height          =   285
      Left            =   720
      TabIndex        =   18
      Top             =   720
      Width           =   675
   End
   Begin VB.CommandButton Command13 
      Caption         =   "出走表get"
      Height          =   525
      Left            =   1650
      TabIndex        =   17
      Top             =   6660
      Width           =   1485
   End
   Begin VB.CommandButton Command12 
      Caption         =   "レースget"
      Height          =   525
      Left            =   90
      TabIndex        =   16
      Top             =   6660
      Width           =   1485
   End
   Begin VB.CommandButton Command11 
      Caption         =   "日get"
      Height          =   525
      Left            =   1650
      TabIndex        =   15
      Top             =   6090
      Width           =   1485
   End
   Begin VB.CommandButton Command10 
      Caption         =   "年get"
      Height          =   525
      Left            =   120
      TabIndex        =   14
      Top             =   6090
      Width           =   1485
   End
   Begin VB.CommandButton Command9 
      Caption         =   "結果get"
      Height          =   525
      Left            =   3150
      TabIndex        =   13
      Top             =   6660
      Width           =   1485
   End
   Begin VB.CommandButton Command8 
      Caption         =   "出馬表get"
      Height          =   525
      Left            =   4080
      TabIndex        =   12
      Top             =   2520
      Width           =   1485
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0000FFFF&
      Caption         =   "南関コンピデータ取得   ]"
      Height          =   525
      Left            =   2790
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   11
      Top             =   870
      Width           =   1485
   End
   Begin VB.TextBox txtUma 
      Height          =   435
      Left            =   1620
      TabIndex        =   9
      Text            =   "4"
      Top             =   5580
      Width           =   1275
   End
   Begin VB.TextBox txtRace 
      Height          =   435
      Left            =   1620
      TabIndex        =   7
      Text            =   "8"
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton Command6 
      Caption         =   "rakuten南関"
      Height          =   525
      Left            =   1650
      TabIndex        =   6
      Top             =   4470
      Width           =   1245
   End
   Begin VB.TextBox txtY 
      Height          =   435
      Left            =   3120
      TabIndex        =   5
      Text            =   "2007"
      Top             =   3780
      Width           =   1275
   End
   Begin VB.CommandButton Command5 
      Caption         =   "yahoo!中央"
      Height          =   525
      Left            =   1680
      TabIndex        =   4
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton Command4 
      Caption         =   "getHTMLString"
      Height          =   525
      Left            =   1680
      TabIndex        =   3
      Top             =   3210
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cmpi"
      Height          =   525
      Left            =   4560
      TabIndex        =   2
      Top             =   1320
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000C0&
      Caption         =   "ierun"
      Height          =   525
      Left            =   120
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "test"
      Height          =   525
      Left            =   4440
      TabIndex        =   0
      Top             =   1920
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Year"
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   68
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "stop"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   21
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "start"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   19
      Top             =   720
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "umaban"
      Height          =   405
      Index           =   1
      Left            =   300
      TabIndex        =   10
      Top             =   5640
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "race No."
      Height          =   405
      Index           =   0
      Left            =   300
      TabIndex        =   8
      Top             =   5100
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const PATH_DB = ".\data\dmc.mdb"

Private objNonCode As Object    ' 文字コード判定/変換オブジェクト
Private strOutCode As String    ' 出力文字コード

Dim mURL As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, _
    ByVal szURL As String, ByVal szFileName As String, _
    ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

' IEを変数として定義
Public WithEvents objIE As InternetExplorer
Attribute objIE.VB_VarHelpID = -1

Public document_completed_flag As Boolean

Private ie As SHDocVw.InternetExplorer
Attribute ie.VB_VarHelpID = -1
Private SaveFileName As String

Private myURL         As String
Private gYmd As String

Private gYmdPlaceRace As String
Private gStr As String
Private gYear() As String
Private gUrlYear() As String
Private gDnmYear() As String
Private gDnmUrlYear() As String
Private gDay() As String
Private gDayFmt() As String
Private gUrlDay() As String
Private gPosDay() As String     '開催場所
Private gPosDayCd() As String     '開催場所コード 楽天
Private gPosDayDbCd() As String     '開催場所コード データベース
Private gCmpDay() As String     'コンピ指数
Private gWk As String
Private gRace() As String
Private gDenmaRace() As String  '出走表
Private gResRace() As String    '結果
Private gUmaban() As String
Private gBamei() As String
Private gUmaCD() As String
Private gCmp() As String
Private gFukuMny() As String
Private gFukuNum() As String

Private gDnmDay() As String
Private gDnmDayFmt() As String
Private gDnmUrlDay() As String
Private gDnmPosDay() As String     '開催場所
Private gDnmPosDayCd() As String     '開催場所コード 楽天
Private gDnmPosDayDbCd() As String     '開催場所コード データベース

Private aBasicDat() As String

Private Function aCheckTime(pSetTime As String, pBeforeMin As Integer) As Variant
    Dim aNow As String
    Dim aHasso As String
    Dim aY As String
    Dim aM As String
    Dim aD As String
    Dim aAns As Variant
    Dim aTime() As String
    
    aHasso = pSetTime '"9:10"
'    aHasso = "11:39"  'for debug
'    pBeforeMin = -2  'for debug
    
    'n分前
    aTime = Split(aHasso, ":")
    aY = Format$(Now, "yyyy")
    aM = Format$(Now, "mm")
    aD = Format$(Now, "dd")
    
    aHasso = aY & "/" & aM & "/" & aD & " " & aTime(0) & ":" & aTime(1) & ":00"
    aAns = DateAdd("n", pBeforeMin, aHasso)   '馬券購入の閾時刻
    Debug.Print aAns
    
    Dim aDiff As Variant
    
    aDiff = DateDiff("s", aAns, Format$(Now, "yyyy/mm/dd hh:nn:ss"))
    
    aCheckTime = aDiff
    
End Function

Private Function ChkTanSan(pArg As String, pUmaban As String, pHimo As String) As Integer
    'https://keiba.yahoo.co.jp/odds/tfw/1705040301/?ninki=1
    myURL = "https://keiba.yahoo.co.jp/odds/tfw/" & pArg & "?ninki=1"
    ie.Navigate2 myURL
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Dim ii As Integer
    Dim jj As Integer
    Dim kk As Integer
    Dim mm As Integer
    Dim str As String
    Dim str2 As String
    Dim aUmaban As String
    Dim aSanrenPuku As String
    Dim aRnk As String
    Dim aTan As String
    Dim aPos As Integer
    Dim aStr As String
    Dim aTargetFlg As Boolean
    Dim aUmas() As String
    Dim aHimo() As String
    Dim aJikuFlg As Boolean
    Dim aHimoCnt As Integer
    
    aStr = getHTMLString(ie)
    
    aDat = Split(aStr, vbLf)
    
    aHimo = Split(pHimo, "-")
    aTargetFlg = False
    
    For ii = 0 To UBound(aDat)
        If aDat(ii) = "<h3 class=""midashi3rd mgnBS"">枠連</h3>" Then
            Exit For
        End If
        aPos = InStr(aDat(ii), "oddsRank")
        If aPos > 0 Then
            aRnk = Mid$(aDat(ii), aPos + 10)
            aPos = InStr(aRnk, "<")
            aRnk = Left$(aRnk, aPos - 1)
            
            aUmaban = aDat(ii + 1)
            aPos = InStr(aUmaban, "</span></td><td>")
            aUmaban = Mid$(aUmaban, aPos + 16)
            aPos = InStr(aUmaban, "<")
            aUmaban = Left$(aUmaban, aPos - 1)
            If Format$(aUmaban, "00") = pUmaban Then
                If Format$(aRnk, "00") > "06" Then
                    aTargetFlg = True
                End If
                Exit For
            End If
        End If
    Next ii
    
    If aTargetFlg = True Then
        '条件に該当していたら、3連複を確認
        'https://keiba.yahoo.co.jp/odds/sf/1705040301/?ninki=1
        myURL = "https://keiba.yahoo.co.jp/odds/sf/" & pArg & "?ninki=1"
        ie.Navigate2 myURL
        Do While ie.Busy = True Or ie.ReadyState <> 4
             DoEvents
        Loop
         
        aStr = getHTMLString(ie)
         
        aDat = Split(aStr, vbLf)
        aJikuFlg = False
        aHimoCnt = 0
        
        For jj = 0 To UBound(aDat)
            aPos = InStr(aDat(jj), "oddsRank")
            If aPos > 0 Then
                aPos = InStr(aDat(jj), "class=""txR"">")    'オッズは、万馬券以上確定
                If aPos > 0 Then
                    aPos = InStr(aDat(jj), "</td><td>")
                    aSanrenPuku = Mid$(aDat(jj), aPos + 9)
                    aPos = InStr(aSanrenPuku, "<")
                    aSanrenPuku = Left$(aSanrenPuku, aPos - 1)
                    aUmas = Split(aSanrenPuku, "－")
                    For kk = 0 To UBound(aUmas)
                        aUmas(kk) = Format$(aUmas(kk), "00")
                    Next kk
                    '軸が含まれているかチェック
                    For kk = 0 To UBound(aUmas)
                        If pUmaban = aUmas(kk) Then
                            aJikuFlg = True
                            Exit For
                        End If
                    Next kk
                    If aJikuFlg = True Then
                        '候補が含まれているかチェック
                        For kk = 0 To UBound(aUmas)
                            For mm = 0 To UBound(aHimo)
                                If aHimo(mm) = aUmas(kk) Then
                                    aHimoCnt = aHimoCnt + 1
                                    Exit For
                                End If
                            Next mm
                        Next kk
                        
                        If aHimoCnt = 2 Then
                            '馬券購入
                            Debug.Print aSanrenPuku
                        End If
                    End If
                    
                End If
            End If
        Next jj
    End If

End Function

'中央競馬の出馬表、結果URLの取得
'<tr>
'           <td class="raceNum leftTD" rowspan=2>1Ｒ</td>
'           <td class="raceName" colspan=2>２歳未勝利(混合)[指定] </td>
'           <td class="raceDist" rowspan=2>ダ1300</td>
'           <td class="racePdf" rowspan=2><p class="btn_etc margT02"><a href="/goku-uma/member/pdf/2016/45/2016110500501N_9.pdf" target=top>東京版PDF&nbsp;<img width="10" height="10" alt="→" src="/goku-uma/img/arrow_green-s.gif"></a></p><p class="btn_etc margT02"><a href="/goku-uma/member/pdf/2016/45/2016110500501NO9.pdf" target=top>大阪版PDF&nbsp;<img width="10" height="10" alt="→" src="/goku-uma/img/arrow_green-s.gif"></a></p><p class="btn_etc margT02"><a href="race.zpl?mode=result&rid=2016110500501"> レース結果 &nbsp;<img width="10" height="10" alt="→" src="/goku-uma/img/arrow_green-s.gif"></a></p></td>
'           <tr>
'           <td class="raceStatus">馬齢</td>
'           <td class="raceDeuma"><p class="btn_etc"><a href="race.zpl?mode=program&rid=2016110500501">出馬表&nbsp;<img width="10" height="10" alt="→" src="/goku-uma/img/arrow_green-s.gif"></a></p></td>
'           </tr>
Sub getDenmaResult()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    Dim aWk As String
    
    '出馬表、結果
    .Pattern = "a href.+レース結果"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        
        If InStr(retstr, "レース結果") = 0 Then
            Exit For
        End If
        
        'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top.zpl?group_id=572&y=2016&wk=43&mode=kako
        cnt = cnt + 1
        ReDim Preserve gDenmaRace(cnt)
        ReDim Preserve gResRace(cnt)
        
        aWk = Mid$(retstr, 9)
        aWk = Left$(aWk, InStr(aWk, "レース結果") - 3)      'レース結果
        
        'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=result&rid=2016110500501
        gResRace(cnt) = Replace("http://p.nikkansports.com" & aWk, "amp;", "")
        
    Next
    
    
    





'    '出馬表、結果
'    .Pattern = "raceNum leftTD.+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+"
'
'    cnt = -1
'    pos = 0
'    Set Matches = .Execute(gStr)   ' 検索を実行します。
'
'    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'        pos = Match.FirstIndex       '一致する文字列が見つかった位置
'        retstr = Match.Value
'
'        If InStr(retstr, "レース結果") = 0 Then
'            Exit For
'        End If
'
'        'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top.zpl?group_id=572&y=2016&wk=43&mode=kako
'        cnt = cnt + 1
'        ReDim Preserve gDenmaRace(cnt)
'        ReDim Preserve gResRace(cnt)
'
'        aWk = Mid$(retstr, InStr(retstr, "race.zpl"))
'        aWk = Left$(aWk, InStr(aWk, "レース結果") - 4)      'レース結果
'
'        'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=result&rid=2016110500501
'        gResRace(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/races/" & aWk, "amp;", "")
'
'        aWk = Mid$(retstr, InStr(retstr, "raceDeuma") + 39)
'        aWk = Left$(aWk, InStr(aWk, "出馬表") - 3)      '出馬表
'
'        'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=program&rid=2016110500501
'        gDenmaRace(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/races/" & aWk, "amp;", "")
'    Next
'
'    'コンピ指数
'    .Pattern = "レース内容.+コンピ指数"
'
'    cnt = -1
'    pos = 0
'    Set Matches2 = .Execute(gStr)   ' 検索を実行します。
'
'    For Each Match In Matches2   ' Matches コレクションに対して繰り返し処理を行います。
'        pos = Match.FirstIndex       '一致する文字列が見つかった位置
'        retstr = Match.Value
'
'
'        aWk = Mid$(retstr, InStr(retstr, "a href=") + 10, 61)
'        aWk = Replace("http://p.nikkansports.com/goku-uma/member" & aWk, "amp;", "")
'
'        ret = URLDownloadToFile(0, aWk, "c:\data\id" & Mid(aWk, 69, 3) & "date" & Mid(aWk, 78, 8) & ".txt", 0, 0)
'        DoEvents
'    Next

End With
    
End Sub

Private Function getHtmlFile() As String
    Dim fnum As Long
    Dim wk As String
    Dim str As String
    
    fnum = FreeFile()
    
    Open SaveFileName & ".txt" For Input As #fnum
    
    Do Until EOF(fnum)
        Line Input #fnum, wk
        str = str & vbLf & wk
    Loop
    
    Close #fnum
    
    getHtmlFile = str
    
End Function

Private Sub getNankanCmpiList()
    Dim str As String
    Dim aYear As Integer
    Dim aDay As Integer
    Dim aRace As Integer
    Dim aUma As Integer
    Dim aGatu As String
    Dim aNiti As String
    Dim aYmd As String
    Dim prt As String
    Dim dbg As String
    Dim timenow As String
    Dim str2 As String
    
    Dim fnum As Integer
    fnum = FreeFile()
    
    timenow = Format$(Now, "hh:mm:ss")
    
'    Open "c:\temp\daily\nankan-" & areaY.Text & areaMD(0).Text & areaMD(1).Text & "-" & Format$(Now, "yyyymmddhhmmss") & ".txt" For Output As #fnum
    
    Me.Caption = "start"
    Me.Refresh
    '指定のURLを表示
    myURL = "http://p.nikkansports.com/goku-uma/member/races/past_list_nankan.zpl"
    
    ie.Navigate2 myURL

    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    str = getHTMLString(ie)

    Me.Caption = "comp"
    Me.Refresh
    
    '年のURLを取得する
    If str = "" Then
        GoTo exit_here
    End If
    gStr = str
    Call getYear(1)
   
    '年ループ   gYear gUrlYear
    For aYear = 2 To UBound(gYear)
        If gYear(aYear) = areaY.Text Or areaY.Text = "" Then
            '指定の年サイトに移動
            myURL = gUrlYear(aYear)
            
            Me.Caption = "start"
            Me.Refresh
            ie.Navigate2 myURL
        
            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop
            str = getHTMLString(ie)
            
            Me.Caption = "comp"
            Me.Refresh
            
            'すべての日付のURLを取得する
            If str = "" Then
                GoTo exit_here
            End If
            gStr = str
            Call getDay(1)
            
            '日付ループ gDay gPosDay
            For aDay = 0 To UBound(gDay)
                If (gDayFmt(aDay) >= areaMD(0).Text And gDayFmt(aDay) <= areaMD(1).Text) Or ("" = areaMD(0).Text And "" = areaMD(1).Text) Then
                    aGatu = Mid$(gDay(aDay), 1, InStr(gDay(aDay), "月") - 1)
                    aNiti = Mid$(gDay(aDay), InStr(gDay(aDay), "月") + 1)
                    aNiti = Left$(aNiti, Len(aNiti) - 1)
                    aYmd = gYear(aYear) & Format$(aGatu, "00") & Format$(aNiti, "00")
                    'If Format$(Now, "yyyymmdd") > aYmd Then
                        '任意の日付
                        
                        Me.Caption = "start"
                        Me.Refresh
                        
                        myURL = gUrlDay(aDay)
                        ie.Navigate2 myURL
'                        ie.Visible = True    'IE を表示
                        Do While ie.Busy = True Or ie.ReadyState <> 4
                            DoEvents
                        Loop
                         
                        Me.Caption = "comp"
                        Me.Refresh
                         
                        '全レースのURLを取得する
                        str = getHTMLString(ie)
                        If str = "" Then
                            GoTo exit_here
                        End If
                        gStr = str
                        Call getRaces
                         
                        'コンピ指数
                        Me.Caption = "start"
                        Me.Refresh
                        
                        myURL = gCmpDay(0)
                        ie.Navigate2 myURL
'                        ie.Visible = True    'IE を表示
                        Do While ie.Busy = True Or ie.ReadyState <> 4
                            DoEvents
                        Loop
                         
                        Me.Caption = "comp"
                        Me.Refresh
                        
                        'コンピ指数 ファイル保存 gCmpDay
                        str = getHTMLString(ie)
                        If str = "" Then
                            GoTo exit_here
                        End If
                        gStr = str
                        
                str2 = Replace(str, vbLf, vbCr & vbLf)
                ff = App.Path & "\" & gYear(aYear) & gDayFmt(aDay) & gPosDay(aDay)
                Call FilePutContents(ff & ".txt", str2, "utf-8")
                        
                        
'                        'レースループ
'                        For aRace = 0 To UBound(gDenmaRace)
'        '                    myURL = gUrlDay(aDay)
'        '                    ie.Navigate2 myURL
'        '                    ie.Visible = True    'IE を表示
'        '                    Do While ie.Busy = True Or ie.ReadyState <> 4
'        '                        DoEvents
'        '                    Loop
'
'                            '出走表
'                            myURL = gDenmaRace(aRace)
'
'                            Me.Caption = "start"
'                            Me.Refresh
'
'                            If chkDL.Value = 0 Then
'                                ie.Navigate2 myURL
'                            Else
'                                ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                                DoEvents
'                            End If
'
'                            If chkDL.Value = 0 Then
'                                Do While ie.Busy = True Or ie.ReadyState <> 4
'                                    DoEvents
'                                Loop
'                                str = getHTMLString(ie)
'                            Else
'                                Call TextCodeChg(SaveFileName)
'                                str = getHtmlFile
'                            End If
'
'                            Me.Caption = "comp"
'                            Me.Refresh
'
''''                            '出走表 すべての馬の馬番と馬名を取得する gBamei gUmaban
''''                            If str = "" Then
''''                                GoTo exit_here
''''                            End If
''''                            gStr = str
''''                            Call getRunTable
''''
''''        '                    myURL = gUrlDay(aDay)
''''        '                    ie.Navigate2 myURL
''''        '                    ie.Visible = True    'IE を表示
''''        '                    Do While ie.Busy = True Or ie.ReadyState <> 4
''''        '                        DoEvents
''''        '                    Loop
''''
''''
''''                            '年月日(gYear(aYear) & gDay(aDay))、開催場所(gPosDay(aDay))、レース番号(gRace(aRace))、馬番、馬名(gBamei gUmaban)をファイルに出力する
''''                            Debug.Print gYear(aYear) & "," & gDay(aDay) & "," & gRace(aRace)
''''                            For aUma = 0 To UBound(gUmaban)
''''                                'Debug.Print gUmaban(aUma) & "," & gBamei(aUma)
''''
''''                                prt = "1," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gUmaban(aUma) & "," & gBamei(aUma) & "," & gCmp(aUma)
''''                                Debug.Print prt
''''                                Print #fnum, prt
''''                            Next aUma
''''
''''                            '結果
''''                            If Format$(Now, "yyyymmdd") > aYmd Then
''''                                If UBound(gResRace) >= aRace Then
''''                                    myURL = gResRace(aRace)
''''
''''                                    Me.Caption = "start"
''''                                    Me.Refresh
''''
''''                                    If chkDL.Value = 0 Then
''''                                        ie.Navigate2 myURL
''''                                    Else
''''                                        ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
''''                                        DoEvents
''''                                    End If
''''
''''                                    If chkDL.Value = 0 Then
''''                                        Do While ie.Busy = True Or ie.ReadyState <> 4
''''                                            DoEvents
''''                                        Loop
''''                                        str = getHTMLString(ie)
''''                                    Else
''''                                        Call TextCodeChg(SaveFileName)
''''                                        str = getHtmlFile
''''                                    End If
''''
''''                                    Me.Caption = "comp"
''''                                    Me.Refresh
''''
''''                                    '結果 とりあえず、複勝のみ gFukuMny gFukuNum
''''                                    If str = "" Then
''''                                        GoTo exit_here
''''                                    End If
''''                                    gStr = str
''''                                    Call getRes
''''
''''                                    '結果(gFukuMny gFukuNum)をファイルに出力する
''''                                    For aUma = 0 To UBound(gFukuNum)
''''                                        Debug.Print gFukuNum(aUma) & "," & gFukuMny(aUma)
''''                                        prt = "2," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gFukuNum(aUma) & "," & gFukuMny(aUma)
''''                                        Debug.Print prt
''''                                        Print #fnum, prt
''''                                    Next aUma
''''                                End If
''''                            End If
''''
''''                            prt = prt
'                        Next aRace
                    'End If
                End If
            Next aDay
        End If
    Next aYear
    
exit_here:
'    Close #fnum
    
    Debug.Print "start:" & timenow
    Debug.Print "end  :" & Format$(Now, "hh:mm:ss")

End Sub

Private Sub old_chuuou()
    Dim dd() As String
  Dim Stream As Object
    Dim str As String
    Dim str2 As String
    Dim kbn As Integer
    
    kbn = 0
    '指定のURLを表示
'    If kbn = 0 Then
''        'コンピ指数
''        myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2016&mode=kako"
'        '結果
'        myURL = "http://p.nikkansports.com/goku-uma/member/result/result_list.zpl?year=2016&mode=kako"
'    Else
'        myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2016&mode=kako"
'    End If
        myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2007&mode=kako"
    
    ie.Navigate2 myURL
    If chkD.Value = 1 Then
        ie.Visible = True    'IE を表示
    End If

    Me.Caption = "Year start"
    Me.Refresh

    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Me.Caption = "Year comp"
    Me.Refresh
    '
    '年のURLを取得する
    str = getHTMLString(ie)
    If str = "" Then
'        GoTo exit_here
    End If
    gStr = str
    Call getYear(kbn)

    '年ループ   gYear gUrlYear
    For aYear = 0 To UBound(gYear)
        If gYear(aYear) = areaY.Text Or areaY.Text = "" Then
            '指定の年サイトに移動
            myURL = gUrlYear(aYear)
            ie.Navigate2 myURL

            If chkD.Value = 1 Then
                ie.Visible = True    'IE を表示
            End If

            Me.Caption = "Year Get start"
            Me.Refresh

            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop

            Me.Caption = "Year Get comp"
            Me.Refresh

            'すべての日付のURLを取得する
            str = getHTMLString(ie)
            If str = "" Then
'                GoTo exit_here
            End If
            gStr = str
            Call getDay(kbn)
            '日付ループ gDay gPosDay
            For aDay = 0 To UBound(gDay)
                If gDayFmt(aDay) >= areaMD(0).Text And gDayFmt(aDay) <= areaMD(1).Text Then
                    myURL = gUrlDay(aDay)
                    ie.Navigate2 myURL
    
                    If chkD.Value = 1 Then
                        ie.Visible = True    'IE を表示
                    End If
    
                    Me.Caption = "Day start"
                    Me.Refresh
    
                    Do While ie.Busy = True Or ie.ReadyState <> 4
                        DoEvents
                    Loop
    
                    Me.Caption = "Day comp"
                    Me.Refresh
    
                    '全レースのURLを取得する +結果
                    str = getHTMLString(ie)
                    If str = "" Then
    '                            GoTo exit_here
                    End If
                    gStr = str
                
                    str2 = Replace(str, vbLf, vbCr & vbLf)
                    ff = App.Path & "\" & gYear(aYear) & gDayFmt(aDay) & gPosDay(aDay)
                    Call FilePutContents(ff & ".txt", str2, "utf-8")
                End If
                
                
                
  
'  ' VB標準のADODB.Streamオブジェクトを作成する
'  Set Stream = CreateObject("ADODB.Stream")
'
'  ' ストリームの文字コードをUTF8に設定する
'  Stream.Charset = "UTF-8"
'  ' ファイルのタイプ(1:バイナリ 2:テキスト)
'  Stream.Type = 2
'  ' ストリームを開く
'  Stream.Open
'  ' ストリームの保存形式をテキスト形式にする
'  Stream.WriteText str
'  ' ストリームに名前を付けて保存する(1は新規作成 2は上書き保存)
'  Stream.SaveToFile (App.Path & "\" & gYear(aYear) & gDayFmt(aDay) & gPosDay(aDay) & ".txt"), 2
'  ' ストリームを閉じる
'  Stream.Close
                
'dd = Split(str, vbLf)
'    For ii = 0 To UBound(dd)
'        str2 = str2 & dd(ii)
'    Next ii
                
'                fnum2 = FreeFile()
'                Open ff & ".txt" For Output As #fnum2
'
'                Print #fnum2, str
'                Close #fnum2
'                ret = URLDownloadToFile(0, myURL, App.Path & "\" & gYear(aYear) & gDayFmt(aDay) & gPosDay(aDay) & "_ex.txt", 0, 0)
            Next aDay
            
'            '日付ループ gDay gPosDay
'            For aDay = 0 To UBound(gDay)
'                If (gDayFmt(aDay) >= areaMD(0).Text And gDayFmt(aDay) <= areaMD(1).Text) Or ("" = areaMD(0).Text And "" = areaMD(1).Text) Then
'                    aGatu = Mid$(gDay(aDay), 1, InStr(gDay(aDay), "月") - 1)
'                    aNiti = Mid$(gDay(aDay), InStr(gDay(aDay), "月") + 1)
'                    aNiti = Left$(aNiti, Len(aNiti) - 1)
'                    aYmd = gYear(aYear) & Format$(aGatu, "00") & Format$(aNiti, "00")
'                    If Format$(Now, "yyyymmdd") > aYmd Then
'                        '任意の日付
'                        myURL = gUrlDay(aDay)
'                        ie.Navigate2 myURL
'
'                        If chkD.Value = 1 Then
'                            ie.Visible = True    'IE を表示
'                        End If
'
'                        Me.Caption = "Day start"
'                        Me.Refresh
'
'                        Do While ie.Busy = True Or ie.ReadyState <> 4
'                            DoEvents
'                        Loop
'
'                        Me.Caption = "Day comp"
'                        Me.Refresh
'
'                        '全レースのURLを取得する +結果
'                        str = getHTMLString(ie)
'                        If str = "" Then
''                            GoTo exit_here
'                        End If
'                        gStr = str
'
'                        Call getChuuouRaces
'
''                        'レースループ
''                        For aRace = 0 To UBound(gRace)
''
''                            '結果
''                            myURL = gResRace(aRace)
''                            ie.Navigate2 myURL
''
''                            If chkD.Value = 1 Then
''                                ie.Visible = True    'IE を表示
''                            End If
''
''                            Me.Caption = "Race start"
''                            Me.Refresh
''
''                            Do While ie.Busy = True Or ie.ReadyState <> 4
''                                DoEvents
''                            Loop
''
''                            Me.Caption = "Race comp"
''                            Me.Refresh
''
''                            '結果 とりあえず、複勝のみ gFukuMny gFukuNum
''                            str = getHTMLString(ie)
''                            If str = "" Then
'''                                GoTo exit_here
''                            End If
''                            gStr = str
''                            Call getRes
''
''                            '年月日(gYear(aYear) & gDay(aDay))、開催場所(gPosDay(aDay))、レース番号(gRace(aRace))、馬番、馬名(gBamei gUmaban)をファイルに出力する
''                            Debug.Print gYear(aYear) & "," & gDay(aDay) & "," & gRace(aRace)
'''                            For aUma = 0 To UBound(gFukuNum)
'''                                Debug.Print gFukuNum(aUma) & "," & gBamei(aUma)
'''                                Print #fnum, "1," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gUmaban(aUma) & "," & gBamei(aUma)
'''                            Next aUma
''
''                            '結果(gFukuMny gFukuNum)をファイルに出力する
''                            For aUma = 0 To UBound(gFukuNum)
''                                Debug.Print gFukuNum(aUma) & "," & gFukuMny(aUma)
''                                Print #fnum, "2," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gFukuNum(aUma) & "," & gFukuMny(aUma)
''                            Next aUma
''                        Next aRace
'                    End If
'                End If
'            Next aDay

        End If
    Next aYear

End Sub

' 指定されたファイルに指定された文字列を出力する
Public Sub FilePutContents(ByVal sFileName As String, sBuffer As String, Optional sEncoding As String, Optional bSaveToWorkbookPath As Boolean)
    Dim oFso As Object
    Dim oFile As Object
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    ' フラグが指定された場合はワークブックのパスに保存する
    If bSaveToWorkbookPath Then
        sFileName = oFso.GetParentFolderName(ActiveWorkbook.FullName) + "\" + sFileName
    End If

    If sEncoding <> "" Then
        ' エンコーディングが指定された場合は ADODB.Stream を利用して文字コードを変換する
        Dim oAdo As Object
        Set oAdo = CreateObject("ADODB.Stream")
        oAdo.Type = 2 'adTypeText
        oAdo.Charset = sEncoding
        
        oAdo.Open
        oAdo.WriteText sBuffer
        
        ' UTF-8 であれば BOM つきで出力されているはずなので削る
        If LCase(sEncoding) = "utf-8" Then
            ' 出力された BOM をスキップして読み込み直す
            oAdo.position = 0   ' Type の変更には Position が 0 である必要あり
            oAdo.Type = 1 'adTypeBinary
            oAdo.position = 3   ' 先頭の 3 bytes（BOM）をスキップ
            Dim sEncodedBuffer As Variant
            sEncodedBuffer = oAdo.Read()
            
            ' ストリームの先頭に戻って内容を再度書きだす
            oAdo.position = 0
            oAdo.Write sEncodedBuffer
            oAdo.SetEos     ' ストリームの最後にゴミが残っているので削る
        End If
        oAdo.SaveToFile (sFileName), 2 'adSaveCreateOverWrite
        oAdo.Close
    Else
        ' エンコーディングが指定されていない場合は FileSystemObject で出力する
        Set oFile = oFso.CreateTextFile(sFileName, True)
        oFile.Write sBuffer
        oFile.Close
    End If
End Sub
Private Sub Command1_Click()
    Dim SaveFileName As String
    Dim DownloadFile As String
    Dim ret          As Long

    SaveFileName = "C:\temp\Test0924.htm"
    DownloadFile = "http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=006&date=20070106"

    ret = URLDownloadToFile(0, DownloadFile, SaveFileName, 0, 0)
    DoEvents

    If ret = 0 Then
        MsgBox "ダウンロードできました。"
    Else
        MsgBox "エラーが発生しました。"
    End If
End Sub

'年get 0:中央, 1:南関
Private Sub getYear(kbn As Integer)
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    If kbn = 0 Then
'         .Pattern = "result_list\.zpl\?year=....&amp;mode=kako"">....年</a>"
         .Pattern = "compi_list\.zpl\?year=....&amp;mode=kako"">....年</a>"
    Else
        'a href="past_list_nankan.zpl?year=2016&mode=kako">2016年</a>
         .Pattern = "past_list_nankan\.zpl\?year=....&amp;mode=kako"">....年</a>"
    End If
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        If firstY = "" Then
            firstY = retstr
        Else
            If firstY = retstr Then
                Exit For
            End If
        End If
        ReDim Preserve gYear(cnt)
        ReDim Preserve gUrlYear(cnt)
        
        gYear(cnt) = Left$(Right$(retstr, 9), 4)
        '南関競馬の結果
        'http://p.nikkansports.com/goku-uma/member/races/past_list_nankan.zpl?year=2015&mode=kako
        '中央競馬のコンピ指数
        'http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2015&mode=kako
        '中央競馬の結果
        'http://p.nikkansports.com/goku-uma/member/result/result_list.zpl?year=2015&mode=kako
        If kbn = 0 Then
'            'コンピ指数
            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/compi/" & Left$(retstr, Len(retstr) - 11)
            '結果
'            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/result/" & Left$(retstr, Len(retstr) - 11)
            gUrlYear(cnt) = Replace(gUrlYear(cnt), "amp;", "")
        Else
            '結果
            gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & Left$(retstr, Len(retstr) - 11)
        End If
        
    Next
    
End With

End Sub

Private Sub getDenmaYear()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    If chkDL.Value = 0 Then
        .Pattern = "pdf_list\.zpl\?y=....&amp;mode=kako"">....年</a>"
    Else
        .Pattern = "pdf_list\.zpl\?y=....&mode=kako"">....年</a>"
    End If
'    .Pattern = "pdf_list\.zpl\?.+"
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        If firstY = "" Then
            firstY = retstr
        Else
            If firstY = retstr Then
                Exit For
            End If
        End If
        ReDim Preserve gDnmYear(cnt)
        ReDim Preserve gDnmUrlYear(cnt)
        
        gDnmYear(cnt) = Left$(Right$(retstr, 9), 4)
        '結果
        gDnmUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & Left$(retstr, Len(retstr) - 11)
        gDnmUrlYear(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/races/" & Left$(retstr, Len(retstr) - 11), "amp;", "")
        
    Next
    
End With

End Sub

Private Sub getPastResYear()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    If chkDL.Value = 0 Then
        .Pattern = "result_list\.zpl\?year=....&amp;mode=kako"">....年</a>"
    Else
'        .Pattern = "pdf_list\.zpl\?y=....&mode=kako"">....年</a>"
        .Pattern = "result_list\.zpl\?year=....&mode=kako"">....年</a>"
    End If
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        If firstY = "" Then
            firstY = retstr
        Else
            If firstY = retstr Then
                Exit For
            End If
        End If
        ReDim Preserve gDnmYear(cnt)
        ReDim Preserve gDnmUrlYear(cnt)
        
        gDnmYear(cnt) = Left$(Right$(retstr, 9), 4)
        '結果 http://p.nikkansports.com/goku-uma/member/result/result_list.zpl?year=2015&mode=kako
'        gDnmUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/result/" & Left$(retstr, Len(retstr) - 11)
        gDnmUrlYear(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/result/" & Left$(retstr, Len(retstr) - 11), "amp;", "")
        
    Next
    
End With

End Sub

Private Sub TextCodeChg(pSrc As String)
    ' テキストをバイト配列で読込
    Dim ipath As String: ipath = pSrc   'App.Path & "\TestUtf8.txt"
    Dim idat() As Byte
    ReDim idat(FileLen(ipath) - 1) As Byte
    Dim intFileNo As Integer
    intFileNo = FreeFile
    Open ipath For Binary As intFileNo
    Get intFileNo, , idat
    Close intFileNo
            
    ' 文字コード判定(blnBin=バイナリ判定無し)
    Dim cod As String: cod = objNonCode.GetCodeName(idat, blnBin:=False)
    
    ' 判定した文字コードをString(UNICODE)に変換
    Dim uni As String
    Select Case cod
        Case "SJIS"
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case "JIS"
            ' JISからUNICODEへの変換
            uni = objNonCode.JIS_To_VbUnicode(idat)
        Case "EUC"
            ' EUCからUNICODEへの変換
            uni = objNonCode.EUC_To_VbUnicode(idat)
        Case "UNICODE"
            ' UNICODEからUNICODEへの変換
            uni = objNonCode.UNICODE_To_VbUnicode(idat)
        Case "UTF7"
            ' UTF-7からUNICODEへの変換
            uni = objNonCode.UTF7_To_VbUnicode(idat)
        Case "UTF8"
            ' UTF-8からUNICODEへの変換
            uni = objNonCode.UTF8_To_VbUnicode(idat)
        Case "BIN"
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case Else
            ' SJISからUNICODEへの変換
            uni = objNonCode.SJIS_To_VbUnicode(idat)
    End Select

    ' 読込ファイルの改行コードをCRLFへ変換
    uni = objNonCode.ChangeReturnToCrLf(uni)
    
    ' String(UNICODE)を出力したい文字コードのByte配列に変換
    Dim odat() As Byte
    cod = "SJIS" 'strOutCode
    Select Case cod
        Case "SJIS"
            ' UNICODEからSJISへの変換
            odat = objNonCode.VbUnicode_To_SJIS(uni)
        Case "JIS"
            ' UNICODEからJISへの変換
            odat = objNonCode.VbUnicode_To_JIS(uni)
        Case "EUC"
            ' UNICODEからEUCへの変換
            odat = objNonCode.VbUnicode_To_EUC(uni)
        Case "UNICODE"
            ' UNICODEからUNICODEへの変換
            odat = objNonCode.VbUnicode_To_UNICODE(uni)
        Case "UTF7"
            ' UNICODEからUTF7への変換
            odat = objNonCode.VbUnicode_To_UTF7(uni)
        Case "UTF8"
            ' UNICODEからUTF8への変換
            odat = objNonCode.VbUnicode_To_UTF8(uni)
        Case Else
            ' UNICODEからSJISへの変換
            odat = objNonCode.VbUnicode_To_SJIS(uni)
    End Select

    ' 出力ファイルをバイナリ形式で出力
    Dim opath As String: opath = pSrc & ".txt" 'App.Path & "\TestOut.txt"
    If Len(Dir(opath)) <> 0 Then
        Kill opath
    End If
    intFileNo = FreeFile
    Open opath For Binary As intFileNo
    Put intFileNo, , odat
    Close intFileNo
End Sub

Private Sub Command10_Click()
    src = "c:\temp\calendar.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

     .Pattern = "past_list_nankan\.zpl\?year=....&mode=kako"">....年</a>"
'     .Pattern = "a href.+"
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        If firstY = "" Then
            firstY = retstr
        Else
            If firstY = retstr Then
                Exit For
            End If
        End If
        ReDim Preserve gYear(cnt)
        ReDim Preserve gUrlYear(cnt)
        
        gYear(cnt) = Left$(Right$(retstr, 9), 4)
        'http://p.nikkansports.com/goku-uma/member/races/past_list_nankan.zpl?year=2015&mode=kako
        gUrlYear(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & Left$(retstr, Len(retstr) - 11)
        
        Debug.Print gUrlYear(cnt)
    Next
    
End With
    
End Sub

Private Sub getDay(kbn As Integer)
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf & "^"
'    kaigyo = vbCr & "$" & vbLf & "^"
'    kaigyo = vbCr & vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    If kbn = 0 Then
'        'コンピ指数
        .Pattern = "<dt>[0-9]+月[0-9]+日.+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf & ".+" & vbLf
        '結果
'        .Pattern = "a href.+[0-9]+月[0-9]+日\("
    Else
        .Pattern = "<dt>[0-9]+月[0-9]+日.+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+nbsp;"
    End If
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        ReDim Preserve gDay(cnt)
        ReDim Preserve gDayFmt(cnt)
        ReDim Preserve gUrlDay(cnt)
        ReDim Preserve gPosDay(cnt)
        ReDim Preserve gPosDayCd(cnt)
        ReDim Preserve gPosDayDbCd(cnt)
        
        gWk = Mid$(retstr, 5)
        gWk = Left$(gWk, InStr(gWk, "<") - 1)

        gDay(cnt) = gWk

        aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
        gDayFmt(cnt) = aWk
        aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
        gDayFmt(cnt) = gDayFmt(cnt) & aWk
        
        If kbn = 0 Then
'            'chuuou
''            '結果
''            gWk = Mid$(retstr, 5)
''            gWk = Left$(gWk, InStr(gWk, "<") - 1)
''
''            gDay(cnt) = gWk
''
''            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
''            gDayFmt(cnt) = aWk
''            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
''            gDayFmt(cnt) = gDayFmt(cnt) & aWk
'
'            gWk = Mid$(retstr, 60)
'            gWk = Left$(gWk, Len(gWk) - 1)
'
'            gDay(cnt) = gWk
'
'            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
'            gDayFmt(cnt) = aWk
'            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
'            gDayFmt(cnt) = gDayFmt(cnt) & aWk
'
'
'            aWk = Mid$(retstr, InStr(retstr, "course_id=") + 33)
'            aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
'            gPosDay(cnt) = aWk

            'http://p.nikkansports.com/goku-uma/member/result/result_day-list.zpl?date=20111211&mode=kako
'            aWk = Mid$(retstr, InStr(retstr, "a href=") + 9, 48)
'            gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/result" & aWk, "amp;", "")

            'コンピ指数
            aWk = retstr 'Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))
            'http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=005&date=20161106
            gUrlDay(cnt) = Mid$(aWk, InStr(aWk, "a href=") + 9, 42)
            gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/compi" & gUrlDay(cnt), "amp;", "")
            aWk = Mid$(aWk, InStr(aWk, "course_id=") + 33)
            aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
            gPosDay(cnt) = aWk
            aWk = Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))
            
            Do
                If InStr(aWk, "course_id=") > 0 Then
                    cnt = cnt + 1

                    ReDim Preserve gDay(cnt)
                    gDay(cnt) = gDay(cnt - 1)
                    ReDim Preserve gDayFmt(cnt)
                    gDayFmt(cnt) = gDayFmt(cnt - 1)

                    ReDim Preserve gUrlDay(cnt)
                    ReDim Preserve gPosDay(cnt)
                    ReDim Preserve gPosDayCd(cnt)
                    ReDim Preserve gPosDayDbCd(cnt)

                    'コンピ指数
                    'http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=005&date=20161106
                    gUrlDay(cnt) = Mid$(aWk, InStr(aWk, "a href=") + 9, 42)
                    gUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/compi" & gUrlDay(cnt), "amp;", "")
                    aWk = Mid$(aWk, InStr(aWk, "course_id=") + 33)
                    aWk = Mid$(aWk, 1, InStr(aWk, "&nbsp") - 1)
                    gPosDay(cnt) = aWk
                    aWk = Mid$(retstr, InStr(retstr, gPosDay(cnt)) + Len(gPosDay(cnt)))

                Else
                   Exit Do
                End If
            Loop
        Else
            'nankan
            gWk = Mid$(retstr, 5)
            gWk = Left$(gWk, InStr(gWk, "<") - 1)
            
            gDay(cnt) = gWk
            
            aWk = Format$(Left$(gDay(cnt), InStr(gDay(cnt), "月") - 1), "00")
            gDayFmt(cnt) = aWk
            aWk = Format$(Mid$(gDay(cnt), InStr(gDay(cnt), "月") + 1, InStr(gDay(cnt), "日") - (InStr(gDay(cnt), "月") + 1)), "00")
            gDayFmt(cnt) = gDayFmt(cnt) & aWk
            
            gWk = Mid$(retstr, InStr(retstr, "/goku-uma"))
            aWk = Mid$(gWk, InStr(gWk, "kako") + 6)
            aWk = Left$(aWk, Len(aWk) - 6)
            gPosDay(cnt) = aWk
            
            Select Case gPosDay(cnt)
            Case "浦和"
                gPosDayCd(cnt) = "18"
                gPosDayDbCd(cnt) = "42"
            Case "船橋"
                gPosDayCd(cnt) = "19"
                gPosDayDbCd(cnt) = "43"
            Case "大井"
                gPosDayCd(cnt) = "20"
                gPosDayDbCd(cnt) = "44"
            Case "川崎"
                gPosDayCd(cnt) = "21"
                gPosDayDbCd(cnt) = "45"
            End Select
            
            gWk = Left$(gWk, InStr(gWk, "kako") + 3)
            'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top_nankan.zpl?date_place_id=244472&mode=kako
            gUrlDay(cnt) = "http://p.nikkansports.com" & gWk
        End If
'        Debug.Print gUrlDay(cnt)
    Next
    
End With
'    Debug.Print "day"

End Sub


Private Sub getDnmDay()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    '出馬表
    .Pattern = "<dt>[0-9]+月[0-9]+日\(" & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        
        'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top.zpl?group_id=572&y=2016&wk=43&mode=kako
'        aWk = Mid$(retstr, InStr(retstr, "a href=") + 9, 48)
        cnt = cnt + 1
        'つづき　レース結果も取得できる
        ReDim Preserve gDnmDay(cnt)
        ReDim Preserve gDnmDayFmt(cnt)
        ReDim Preserve gDnmUrlDay(cnt)
        ReDim Preserve gDnmPosDay(cnt)
        ReDim Preserve gDnmPosDayCd(cnt)
        ReDim Preserve gDnmPosDayDbCd(cnt)
        
        gWk = Mid$(retstr, 4)
        gWk = Left$(gWk, InStr(gWk, "(") - 1)
        
        gDnmDay(cnt) = gWk
        
        aWk = Format$(Left$(gDnmDay(cnt), InStr(gDnmDay(cnt), "月") - 1), "00")
        gDnmDayFmt(cnt) = aWk
        aWk = Format$(Mid$(gDnmDay(cnt), InStr(gDnmDay(cnt), "月") + 1, InStr(gDnmDay(cnt), "日") - (InStr(gDnmDay(cnt), "月") + 1)), "00")
        gDnmDayFmt(cnt) = gDnmDayFmt(cnt) & aWk
        
        aWk = Mid$(retstr, InStr(retstr, "href") + 7)
        aWk = Left$(aWk, InStr(aWk, "kako") + 3)
        gDnmUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/races" & aWk, "amp;", "")
    Next
    
End With
'    Debug.Print "Dnmday"

End Sub

Private Sub getResDay()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    '出馬表
    .Pattern = "a href.+\(.+"
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        
        'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top.zpl?group_id=572&y=2016&wk=43&mode=kako
'        aWk = Mid$(retstr, InStr(retstr, "a href=") + 9, 48)
        cnt = cnt + 1
        'つづき　レース結果も取得できる
        ReDim Preserve gDnmDay(cnt)
        ReDim Preserve gDnmDayFmt(cnt)
        ReDim Preserve gDnmUrlDay(cnt)
        ReDim Preserve gDnmPosDay(cnt)
        ReDim Preserve gDnmPosDayCd(cnt)
        ReDim Preserve gDnmPosDayDbCd(cnt)
        
        gWk = Mid$(retstr, 60)
        gWk = Left$(gWk, InStr(gWk, "(") - 1)
        
        gDnmDay(cnt) = gWk
        
        aWk = Format$(Left$(gDnmDay(cnt), InStr(gDnmDay(cnt), "月") - 1), "00")
        gDnmDayFmt(cnt) = aWk
        aWk = Format$(Mid$(gDnmDay(cnt), InStr(gDnmDay(cnt), "月") + 1, InStr(gDnmDay(cnt), "日") - (InStr(gDnmDay(cnt), "月") + 1)), "00")
        gDnmDayFmt(cnt) = gDnmDayFmt(cnt) & aWk
        
        aWk = Mid$(retstr, InStr(retstr, "href") + 7)
        aWk = Left$(aWk, InStr(aWk, "kako") + 3)
        gDnmUrlDay(cnt) = Replace("http://p.nikkansports.com/goku-uma/member/result" & aWk, "amp;", "")
    Next
    
End With
'    Debug.Print "Dnmday"

End Sub

'日get
Private Sub Command11_Click()
    src = "c:\temp\calendar.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

     .Pattern = "<dt>[0-9]+月[0-9]+日.+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+nbsp;"
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        ReDim Preserve gDay(cnt)
        ReDim Preserve gUrlDay(cnt)
        ReDim Preserve gPosDay(cnt)
        
        gWk = Mid$(retstr, 5)
        gWk = Left$(gWk, InStr(gWk, "<") - 1)
        
        gDay(cnt) = gWk
        gWk = Mid$(retstr, InStr(retstr, "/goku-uma"))
        aWk = Mid$(gWk, InStr(gWk, "kako") + 6)
        aWk = Left$(aWk, Len(aWk) - 6)
        gPosDay(cnt) = aWk
        gWk = Left$(gWk, InStr(gWk, "kako") + 3)
        'http://p.nikkansports.com/goku-uma/member/races/pdf_list_top_nankan.zpl?date_place_id=244472&mode=kako
        gUrlDay(cnt) = "http://p.nikkansports.com" & gWk
        
'        Debug.Print gUrlDay(cnt)
    Next
    
End With
    Debug.Print "finish"
    
    
End Sub

'レースget
'中央競馬 2011年12月25日 中山 1R
'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=result&rid=2011122500601
'中央競馬 2011年12月25日 中山 11R
'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=result&rid=2011122500611
'中央競馬 2011年12月25日 阪神 4R
'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=result&rid=2011122500904
Private Sub getRaces()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
'    kaigyo = vbCr & "$" & vbLf & "^"
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    'レース数取得
'    .Pattern = "raceNum leftTD""\>.+\</td\>"
'    .Pattern = "raceNum leftTD.+\>.+\</td\>" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+コンピ.+a href=.+"
'    .Pattern = "raceNum leftTD.+\>.+\</td\>" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & "" & kaigyo & ".+" & kaigyo & ".+コンピ.+a href=.+"
    .Pattern = "raceNum leftTD.+\>.+\</td\>" & kaigyo
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        
        ReDim Preserve gRace(cnt)

        gWk = Mid$(retstr, 29)
        gWk = Left$(gWk, 2)
        If Right$(gWk, 1) = "Ｒ" Then
            gWk = Left$(gWk, 1)
        End If

        gRace(cnt) = gWk
        
    Next
    
    '出走表取得
    .Pattern = "href="".+""\>出走表"
'    .Pattern = ".+出走表"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        ReDim Preserve gDenmaRace(cnt)

        gWk = Mid$(retstr, 7)
        gWk = Left$(gWk, Len(gWk) - 5)
        'http://p.nikkansports.com/goku-uma/member/races/race_nankan.zpl?date_place_id=247006&race_id=78400&vw=de&mode=kako
        gDenmaRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
        gDenmaRace(cnt) = Replace(gDenmaRace(cnt), "amp;", "")
        
    Next
    
    '結果URL取得
'    .Pattern = "コンピ&nbsp;\<img width=""10"" height=""10"" alt=""→"" src=""/goku-uma/img/arrow_orange-s.gif""\>\</a\>\</p\>\<p class=""btn_etc margT02""\>\<a href="".+""\>結　果"
    .Pattern = "コンピ&nbsp;\<img width=""10"" height=""10"" alt=""→"" src=""/goku-uma/img/arrow_orange-s.gif""\>\</a\>\</p\>\<p class=""btn_etc margT02""\>\<a href="".+""\>結　果"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        ReDim Preserve gResRace(cnt)

        gWk = Mid$(retstr, 129)
        gWk = Left$(gWk, Len(gWk) - 5)
        
        gResRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
        gResRace(cnt) = Replace(gResRace(cnt), "amp;", "")
        
    Next
    
    'コンピ指数URL取得
    .Pattern = "href="".+""\>コンピ指数"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        ReDim Preserve gCmpDay(cnt)

        gWk = Mid$(retstr, 7)
        gWk = Left$(gWk, Len(gWk) - 7)
        'http://p.nikkansports.com/goku-uma/member/compi/compi_nankan.zpl?date_place_id=247006
        gCmpDay(cnt) = "http://p.nikkansports.com" & gWk
        
    Next
    
    
End With
'    Debug.Print "finish"

End Sub

Private Sub getChuuouRaces()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
'    kaigyo = vbCr & "$" & vbLf & "^"
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    Dim aWk As String
    
    '開催場所取得
    .Pattern = "id=""course.+/span"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        ReDim Preserve gPosDay(cnt)
        ReDim Preserve gPosDayCd(cnt)
        ReDim Preserve gPosDayDbCd(cnt)

        gWk = Mid$(retstr, InStr(retstr, "日") + 2)
        gWk = Left$(gWk, InStr(gWk, "<") - 1)

        gPosDay(cnt) = gWk
        
    Next
    
    
    
    
    
    
    'レース数取得
'    .Pattern = "raceNum leftTD""\>.+\</td\>"
'    .Pattern = "raceNum leftTD.+\>.+\</td\>" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+コンピ.+a href=.+"
    .Pattern = "raceNum leftTD.+\>.+\</td\>" & kaigyo & ".+" & kaigyo & ".+結果"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        
        ReDim Preserve gRace(cnt)
        ReDim Preserve gResRace(cnt)

        gWk = Mid$(retstr, 29)
        gWk = Left$(gWk, 2)
        If Right$(gWk, 1) = "Ｒ" Then
            gWk = Left$(gWk, 1)
        End If

        gRace(cnt) = gWk
        
        '結果URL取得
        gWk = Mid$(retstr, InStr(retstr, "a href=") + 8)
        gWk = Left$(gWk, InStr(gWk, "結果") - 3)
        gWk = Replace(gWk, "amp;", "")
        
        gResRace(cnt) = "http://p.nikkansports.com" & gWk
        
    Next
    
'    '出走表取得
'    .Pattern = "href="".+""\>出走表"
''    .Pattern = ".+出走表"
'
'    cnt = -1
'    pos = 0
'    Set Matches = .Execute(gStr)   ' 検索を実行します。
'
'    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'        pos = Match.FirstIndex       '一致する文字列が見つかった位置
'        retstr = Match.Value
'        cnt = cnt + 1
'
'        ReDim Preserve gDenmaRace(cnt)
'
'        gWk = Mid$(retstr, 7)
'        gWk = Left$(gWk, Len(gWk) - 5)
'        'http://p.nikkansports.com/goku-uma/member/races/race_nankan.zpl?date_place_id=247006&race_id=78400&vw=de&mode=kako
'        gDenmaRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
'        gDenmaRace(cnt) = Replace(gDenmaRace(cnt), "amp;", "")
'
'    Next
    
'    '結果URL取得
''    .Pattern = "コンピ&nbsp;\<img width=""10"" height=""10"" alt=""→"" src=""/goku-uma/img/arrow_orange-s.gif""\>\</a\>\</p\>\<p class=""btn_etc margT02""\>\<a href="".+""\>結　果"
'    .Pattern = "コンピ&nbsp;\<img width=""10"" height=""10"" alt=""→"" src=""/goku-uma/img/arrow_orange-s.gif""\>\</a\>\</p\>\<p class=""btn_etc margT02""\>\<a href="".+""\>結　果"
'
'    cnt = -1
'    pos = 0
'    Set Matches = .Execute(gStr)   ' 検索を実行します。
'
'    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'        pos = Match.FirstIndex       '一致する文字列が見つかった位置
'        retstr = Match.Value
'        cnt = cnt + 1
'
''        Debug.Print retstr
'        ReDim Preserve gResRace(cnt)
'
'        gWk = Mid$(retstr, 129)
'        gWk = Left$(gWk, Len(gWk) - 5)
'
'        gResRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
'        gResRace(cnt) = Replace(gResRace(cnt), "amp;", "")
'
'    Next
'
'    'コンピ指数URL取得
'    .Pattern = "href="".+""\>コンピ指数"
'
'    cnt = -1
'    pos = 0
'    Set Matches = .Execute(gStr)   ' 検索を実行します。
'
'    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
'        pos = Match.FirstIndex       '一致する文字列が見つかった位置
'        retstr = Match.Value
'        cnt = cnt + 1
'
''        Debug.Print retstr
'        ReDim Preserve gCmpDay(cnt)
'
'        gWk = Mid$(retstr, 7)
'        gWk = Left$(gWk, Len(gWk) - 7)
'        'http://p.nikkansports.com/goku-uma/member/compi/compi_nankan.zpl?date_place_id=247006
'        gCmpDay(cnt) = "http://p.nikkansports.com" & gWk
'
'    Next
    
    
End With
'    Debug.Print "finish"

End Sub

'レースget
Private Sub Command12_Click()
    src = "c:\temp\races.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    'レース数取得
    .Pattern = "raceNum leftTD""\>.+\</td\>"
    
    Dim aWk As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        Debug.Print retstr
        
        ReDim Preserve gRace(cnt)

        gWk = Mid$(retstr, 17)
        gWk = Left$(gWk, Len(gWk) - 6)

        gRace(cnt) = gWk
        
    Next
    
    '出走表取得
    .Pattern = "href="".+""\>出走表"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
        ReDim Preserve gDenmaRace(cnt)

        gWk = Mid$(retstr, 7)
        gWk = Left$(gWk, Len(gWk) - 5)
        'http://p.nikkansports.com/goku-uma/member/races/race_nankan.zpl?date_place_id=247006&race_id=78400&vw=de&mode=kako
        gDenmaRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
        
        Debug.Print gDenmaRace(cnt)
    Next
    
    '結果URL取得
    .Pattern = "コンピ&nbsp;\<img width=""10"" height=""10"" src=""/goku-uma/img/arrow_orange-s.gif"" alt=""→""\>\</a\>\</p\>\<p class=""btn_etc margT02""\>\<a href="".+""\>結　果"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        ReDim Preserve gDenmaRace(cnt)

        gWk = Mid$(retstr, 129)
        gWk = Left$(gWk, Len(gWk) - 5)
        
        gDenmaRace(cnt) = "http://p.nikkansports.com/goku-uma/member/races/" & gWk
        
        Debug.Print gDenmaRace(cnt)
    Next
    
    'コンピ指数URL取得
    .Pattern = "href="".+""\>コンピ指数"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        ReDim Preserve gResRace(cnt)

        gWk = Mid$(retstr, 7)
        gWk = Left$(gWk, Len(gWk) - 7)
        'http://p.nikkansports.com/goku-uma/member/compi/compi_nankan.zpl?date_place_id=247006
        gResRace(cnt) = "http://p.nikkansports.com" & gWk
        
        Debug.Print gResRace(cnt)
    Next
    
    
End With
    Debug.Print "finish"

End Sub

'出走表get
Private Sub getRunTable()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
'    kaigyo = vbCr & "$" & vbLf & "^"
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

     .Pattern = "馬番 -->" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        
        ReDim Preserve gUmaban(cnt)
        ReDim Preserve gBamei(cnt)
        ReDim Preserve gCmp(cnt)
        
        If chkDL.Value = 0 Then
            gWk = Mid$(retstr, 12)
            gWk = Left$(gWk, InStr(gWk, "/td") - 2)
            gUmaban(cnt) = gWk
            gWk = Mid$(retstr, InStr(retstr, "horseName2") + 16)
            gWk = Left$(gWk, Len(gWk) - 67)
            gBamei(cnt) = gWk
            gWk = Left$(Mid$(retstr, InStr(retstr, "ftR") + 6), 2)
            gCmp(cnt) = gWk
        Else
            gWk = Mid$(retstr, 21)
            gWk = Left$(gWk, InStr(gWk, "/td") - 2)
            gUmaban(cnt) = gWk
            gWk = Mid$(retstr, InStr(retstr, "horseName2") + 16)
            gWk = Left$(gWk, InStr(gWk, "/span") - 2)
            
            If gWk = "グラッツェーラ" Then
                gWk = gWk
            End If
            
            gBamei(cnt) = gWk
            gWk = Left$(Mid$(retstr, InStr(retstr, "ftR") + 6), 2)
            If gWk = "<f" Then
                gWk = Left$(Mid$(retstr, InStr(retstr, "ftR") + 22), 2)
            End If
            gCmp(cnt) = gWk
        End If
        
        
'        Debug.Print gWk
    Next
    
End With

End Sub

'出走表get　中央競馬
Private Sub getChuuouRunTable()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    Dim aWk As String
    Dim aPos As Integer
    
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    .Pattern = "td>.+</td>" & kaigyo & ".+<td class=""horse"">.+</td>"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        
        aWk = retstr
        
        '馬番
        cnt = cnt + 1
        
        ReDim Preserve gUmaban(cnt)
        ReDim Preserve gBamei(cnt)
        
        aPos = InStr(aWk, "</td>")
        gWk = Left$(aWk, aPos - 1)
        gWk = Mid$(gWk, 4)
        
        gUmaban(cnt) = gWk
        
        '馬名
        gWk = Mid$(aWk, InStr(aWk, "horse") + 7)
        gWk = Left$(gWk, InStr(gWk, "<") - 1)
        gBamei(cnt) = gWk
        
    Next
    
    'gFukuNum(Dnm) & "," & gFukuMny(Dnm)
    .Pattern = "複勝.+円"
    
    cnt = -1
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        
        aWk = retstr
        
        '複勝の馬番
        cnt = cnt + 1
        
        ReDim Preserve gFukuNum(cnt)
        ReDim Preserve gFukuMny(cnt)
        
        gWk = Mid$(aWk, 12)
        aPos = InStr(gWk, "</td>")
        gWk = Left$(gWk, aPos - 1)
        
        gFukuNum(cnt) = gWk
        
        '払い戻し
        aPos = InStr(aWk, gFukuNum(cnt))
        gWk = Mid$(aWk, aPos + Len(gFukuNum(cnt)))
'        aPos = InStr(aWk, "<td>")
        gWk = Mid$(gWk, 10)
        gFukuMny(cnt) = Left$(gWk, Len(gWk) - 1)
        gFukuMny(cnt) = Replace(gFukuMny(cnt), ",", "")
        
        
    Next
    
End With

End Sub

'出走表get
Private Sub Command13_Click()
    src = "c:\temp\denma.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

     .Pattern = "馬番 -->" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo & ".+" & kaigyo
    
    Dim firstY As String
    cnt = -1
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
        pos = Match.FirstIndex       '一致する文字列が見つかった位置
        retstr = Match.Value
        cnt = cnt + 1
        
'        Debug.Print retstr
        
        ReDim Preserve gUmaban(cnt)
        ReDim Preserve gBamei(cnt)
        
        gWk = Mid$(retstr, 22)
        gWk = Left$(gWk, InStr(gWk, "/td") - 2)
        gUmaban(cnt) = gWk
        Debug.Print gWk
        gWk = Mid$(retstr, InStr(retstr, "horseName2") + 16)
        gWk = Left$(gWk, Len(gWk) - 9)
        gBamei(cnt) = gWk
        
        Debug.Print gWk
    Next
    
End With

End Sub

'
'
'中央競馬
Private Sub Command14_Click()
    
    Dim str As String
    Dim aYear As Integer
    Dim aDay As Integer
    Dim aRace As Integer
    Dim aUma As Integer
    Dim aGatu As String
    Dim aNiti As String
    Dim aYmd As String
    Dim kbn As Integer
    Dim aSlt() As String
    Dim ii As Long
    Dim ff As String
    
    kbn = 0
    
    Dim fnum As Integer
    Dim fnum2 As Integer
    fnum = FreeFile()
    
    Open "c:\temp\chuuou" & areaY.Text & areaMD(0).Text & areaMD(1).Text & "-" & Format$(Now, "yyyymmddhhmmss") & ".txt" For Output As #fnum
    
    GoTo Dnm
    
    
Dnm:
    
Dim ret          As Long

'DownloadFile = "http://p.nikkansports.com/goku-uma/member/races/pdf_list.zpl?y=2016&wk=45&mode=kako"
'
'ret = URLDownloadToFile(0, DownloadFile, SaveFileName, 0, 0)
'DoEvents
'
'If ret = 0 Then
'    MsgBox "ダウンロードできました。"
'Else
'    MsgBox "エラーが発生しました。"
'End If
    
    
    '出馬表
    'http://p.nikkansports.com/goku-uma/member/races/pdf_list.zpl?y=2016&wk=45&mode=kako
    If optMode(0).Value = True Then
        myURL = "http://p.nikkansports.com/goku-uma/member/result/result_list.zpl?year=2016&mode=kako"  'レース結果 2007年以降のデータ
    Else
        myURL = "http://p.nikkansports.com/goku-uma/member/races/pdf_list.zpl?y=2016&wk=45&mode=kako"   '過去のPDF出馬表 2011年以降のデータ
    End If
    
    If chkDL.Value = 0 Then
        ie.Navigate2 myURL
    Else
        ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
        DoEvents
    End If
    
    If chkD.Value = 1 And chkDL.Value = 0 Then
        ie.Visible = True    'IE を表示
    End If
    
    Me.Caption = "Year start"
    Me.Refresh
    
    If chkDL.Value = 0 Then
        Do While ie.Busy = True Or ie.ReadyState <> 4
            DoEvents
        Loop
        '年のURLを取得する
        str = getHTMLString(ie)
    Else
        Call TextCodeChg(SaveFileName)
        str = getHtmlFile
    End If
    
    Me.Caption = "Year comp"
    Me.Refresh
    
    If str = "" Then
        GoTo exit_here
    End If
        
    gStr = str
    
    If optMode(0).Value = True Then
        Call getPastResYear         'つづき
    Else
        Call getDenmaYear
    End If
    
    '年ループ   gYear gUrlYear
    For aYear = 0 To UBound(gDnmYear)
        If gDnmYear(aYear) = areaY.Text Or areaY.Text = "" Then
            '指定の年サイトに移動
            myURL = gDnmUrlYear(aYear)
'            If chkDL.Value = 0 Then
                ie.Navigate2 myURL
'            Else
'                ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                DoEvents
'            End If
            
'            If chkD.Value = 1 And chkDL.Value = 0 Then
'                ie.Visible = True    'IE を表示
'            End If
'
            Me.Caption = "Year Get start"
            Me.Refresh
            
'            If chkDL.Value = 0 Then
                Do While ie.Busy = True Or ie.ReadyState <> 4
                    DoEvents
                Loop
                str = getHTMLString(ie)
'            Else
'                Call TextCodeChg(SaveFileName)
'                str = getHtmlFile
'            End If
            
            Me.Caption = "Year Get comp"
            Me.Refresh
            
            'すべての日付のURLを取得する
            If str = "" Then
                GoTo exit_here
            End If
            gStr = str
            
            If optMode(0).Value = True Then
                Call getResDay
            Else
                Call getDay(0)
            End If
            
            '日付ループ gDay gPosDay
            For aDay = 0 To UBound(gDnmDay)
                If (gDnmDayFmt(aDay) >= areaMD(0).Text And gDnmDayFmt(aDay) <= areaMD(1).Text) Or ("" = areaMD(0).Text And "" = areaMD(1).Text) Then
                    aGatu = Mid$(gDnmDay(aDay), 1, InStr(gDnmDay(aDay), "月") - 1)
                    aNiti = Mid$(gDnmDay(aDay), InStr(gDnmDay(aDay), "月") + 1)
                    aNiti = Left$(aNiti, Len(aNiti) - 1)
                    aYmd = gDnmYear(aYear) & Format$(aGatu, "00") & Format$(aNiti, "00")
                    If Format$(Now, "yyyymmdd") > aYmd Then
                        
                        myURL = gDnmUrlDay(aDay)
                        'If chkDL.Value = 0 Then
                            ie.Navigate2 myURL
'                        Else
'                            ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                            DoEvents
'                        End If
                        
'                        If chkD.Value = 1 And chkDL.Value = 0 Then
'                            ie.Visible = True    'IE を表示
'                        End If
                        
                        Me.Caption = "Day start"
                        Me.Refresh
                        
'                        If chkDL.Value = 0 Then
                            Do While ie.Busy = True Or ie.ReadyState <> 4
                                DoEvents
                            Loop
                            str = getHTMLString(ie)
'                        Else
'                            Call TextCodeChg(SaveFileName)
'                            str = getHtmlFile
'                        End If
                        
                        Me.Caption = "Day comp"
                        Me.Refresh
                        
                        'すべての日付のURLを取得する
                        If str = "" Then
                            GoTo exit_here
                        End If
                        gStr = str
                        
                        '出馬表、結果URLの取得
                        Call getDenmaResult
                        
                        For aRace = 0 To UBound(gDenmaRace)
                            '出馬表
                            'http://p.nikkansports.com/goku-uma/member/races/race.zpl?mode=program&rid=2016110500801
                            '                                                                          yyyymmdd開催場所,レース番号
                            myURL = gResRace(aRace)
                            gYmdPlaceRace = Right$(myURL, 13)
                            
                            If chkDL.Value = 0 Then
                                ie.Navigate2 myURL
                            Else
                                ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
                                'copy
                                FileCopy SaveFileName, gYmdPlaceRace & "-denma-" & Format$(Now, "yyyymmddhhnnss") & ".txt"
                                DoEvents
                            End If
                            
'                            If chkD.Value = 1 And chkDL.Value = 0 Then
'                                ie.Visible = True    'IE を表示
'                            End If
                            
                            Me.Caption = "Denma start"
                            Me.Refresh
                            
                            If chkDL.Value = 0 Then
                                Do While ie.Busy = True Or ie.ReadyState <> 4
                                    DoEvents
                                Loop
                                str = getHTMLString(ie)
                                
                                fnum2 = FreeFile()
                                ff = App.Path & "\" & gYmdPlaceRace
                                Open ff & ".txt" For Output As #fnum2
                                
                                Print #fnum2, str
                                Close #fnum2
                                
'                                Call TextCodeChg(ff & ".txt")
                            Else
                                Call TextCodeChg(SaveFileName)
                                str = getHtmlFile
                            End If
                            
                            Me.Caption = "Denma comp"
                            Me.Refresh
                            
                            'すべての日付のURLを取得する
                            If str = "" Then
                                GoTo exit_here
                            End If
                            gStr = str
                            
                            Call getChuuouRunTable
                            
                            For Dnm = 0 To UBound(gBamei)
                                prt = "1," & gYmdPlaceRace & "," & gUmaban(Dnm) & "," & gBamei(Dnm)
                                Debug.Print prt
                                Print #fnum, prt
                            Next Dnm
                            
'                            '結果
'                            myURL = gResRace(aRace)
'                            If chkDL.Value = 0 Then
'                                ie.Navigate2 myURL
'                            Else
'                                ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                                'copy
'                                FileCopy SaveFileName, gYmdPlaceRace & "-reslt-" & Format$(Now, "yyyymmddhhnnss") & ".txt"
'                                DoEvents
'                            End If
'
''                            If chkD.Value = 1 And chkDL.Value = 0 Then
''                                ie.Visible = True    'IE を表示
''                            End If
'
'                            Me.Caption = "Result start"
'                            Me.Refresh
'
'                            If chkDL.Value = 0 Then
'                                Do While ie.Busy = True Or ie.ReadyState <> 4
'                                    DoEvents
'                                Loop
'                                str = getHTMLString(ie)
'                            Else
'                                Call TextCodeChg(SaveFileName)
'                                str = getHtmlFile
'                            End If
'
'                            Me.Caption = "Result comp"
'                            Me.Refresh
'
'                            'すべての日付のURLを取得する
'                            If str = "" Then
'                                GoTo exit_here
'                            End If
'                            gStr = str
'
'                            Call getChuuouRes
                            
                            For Dnm = 0 To UBound(gFukuNum)
                                prt = "2," & gYmdPlaceRace & "," & gFukuNum(Dnm) & "," & gFukuMny(Dnm)
                                Debug.Print prt
                                Print #fnum, prt
                            Next Dnm
                            
                        Next aRace
                    End If
                    
                End If
            Next aDay
        End If
    Next aYear
    
exit_here:
    
    Close #fnum

End Sub

Private Sub Command15_Click()
    areaMD(0).Text = ""
    areaMD(1).Text = ""
    areaY.Text = ""

End Sub

Private Sub Command16_Click()
    src = "C:\temp\cmpi\cmpiList.txt"
    SaveFileName = "C:\temp\cmpi\Raw\"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    cnt = 1611
    
    Do Until EOF(fn)
        Line Input #fn, wk
        ret = URLDownloadToFile(0, wk, SaveFileName & Format$(cnt, "0000") & ".txt", 0, 0)
        cnt = cnt + 1
        DoEvents
    Loop
    
    '<<ファイル 閉>>
    Close #fn

    MsgBox "finish"
End Sub

Private Sub Command18_Click()
    Dim StartTime  As Long
    Dim StopTime  As Long

    StartTime = GetTickCount
   
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   
    Dim str As String
    
   
   '指定のURLを表示
    myURL = "http://keiba.rakuten.co.jp/?l-id=top_headernavi_1st_top/"
   objIE.Navigate2 myURL
   objIE.Visible = True    'IE を表示
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
    Loop
    document_completed_flag = False

    Sleep (100)

'    str = getHTMLString(objIE)

    myURL = "http://keiba.rakuten.co.jp/race_card/list/RACEID/201702101914110500"
   objIE.Navigate2 myURL
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
    Loop
    document_completed_flag = False

    Sleep (100)

'    str = getHTMLString(objIE)

    myURL = "http://keiba.rakuten.co.jp/odds/tanfuku/RACEID/201702101914110501"
   objIE.Navigate2 myURL
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
    Loop
    document_completed_flag = False
    
    Sleep (100)
    
'    str = getHTMLString(ie)

'    myURL = "http://keiba.rakuten.co.jp/race_card/list/RACEID/201702101914110500"
'   objIE.Navigate2 myURL
'    Do While objIE.Busy = True Or objIE.ReadyState <> 4
'        DoEvents
'        If document_completed_flag = True Then
'            Exit Do
'        End If
'    Loop
'    document_completed_flag = False
'
'    Sleep (100)

    myURL = "http://keiba.rakuten.co.jp/odds/sanrenfuku/RACEID/201702101914110501#headline"
'    myURL = "http://keiba.rakuten.co.jp/odds/sanrentan/RACEID/201702101914110501#headline"
   objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
        Sleep (100)
'        StopTime = GetTickCount
'        If StopTime - StartTime > 5000 Then
'            str = getHTMLString(objIE)
'            If InStr(str, "</html>") > 0 Then
'                Exit Do
'            End If
'        End If
    Loop
    document_completed_flag = False
    
    Dim fn2  As Long
    fn2 = FreeFile
    Open "c:\0210.html" For Output As #fn2
    Print #fn2, str
    
    Close #fn2
'    Exit Sub
    
    Sleep (100)

'    str = getHTMLString(objIE)

    Dim i As Integer
    Dim sel As Object
    
   objIE.Visible = True    'IE を表示
'    objIE.Document.All("selectedJiku")(4).Selected = True
    
    Set sel = objIE.Document.getElementsByName("selectedJiku")(0)
'    sel.selectedIndex = 4
    For i = 0 To sel.length - 1 'select内のタグで
        If sel(i).Value = "4" Then '100のがあったら
            sel(i).Selected = True '選ぶ
            Exit For '選んだから終わる
        End If
    Next i
    
   objIE.Navigate2 "JavaScript:displayOdds()"
    
    

End Sub

Private Sub Command19_Click()
    Dim StartTime  As Long
    Dim StopTime  As Long
   
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   
    Dim str As String

    myURL = "http://keiba.rakuten.co.jp/odds/sanrenfuku/RACEID/201702101914110501#headline"
'    myURL = "http://keiba.rakuten.co.jp/odds/sanrentan/RACEID/201702101914110501#headline"
   objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
        Sleep (100)
        StopTime = GetTickCount
        If StopTime - StartTime > 5000 Then
            str = getHTMLString(objIE)
            If InStr(str, "</html>") > 0 Then
                Exit Do
            End If
        End If
    Loop
    document_completed_flag = False
    
'    Exit Sub
    
    Sleep (100)

'    str = getHTMLString(objIE)

    Dim i As Integer
    Dim sel As Object
    
   objIE.Visible = True    'IE を表示
'    objIE.Document.All("selectedJiku")(4).Selected = True
    
    Set sel = objIE.Document.getElementsByName("selectedJiku")(0)
'    sel.selectedIndex = 4
    For i = 0 To sel.length - 1 'select内のタグで
        If sel(i).Value = "4" Then '100のがあったら
            sel(i).Selected = True '選ぶ
            Exit For '選んだから終わる
        End If
    Next i
    
   objIE.Navigate2 "JavaScript:displayOdds()"
End Sub

Private Sub Command20_Click()
    Dim fn2  As Long
    Dim StartTime  As Long
    Dim StopTime  As Long
   Dim a As Object
   Dim b As Object
   Dim d As Object
   Dim c As String
   
    Set a = CreateObject("Shell.Application")
    Set b = a.Windows()
   
   'ログインしているブラウザを操作する(jun@buhi-buhi.com 名前)
    For Each d In b
        c = ""
        c = d.LocationURL
        If c = "http://www.netkeiba.com/?acc_param=top" Then
            Set objIE = d
            Exit For
        End If
    Next
   
    Sleep (100)
    DoEvents
    
    Dim str As String
    
    'http://nar.netkeiba.com/?pid=schedule&year=2015&month=8
    'スケジュールから、全URLを取得する
    'リンクは、本日以降しかないので、自力で作成する
    'http://nar.netkeiba.com/?pid=race&id=p201747022201&mode=top
    'スケジュールの印から、レースの存在を確認する
    '１Rは、確実に存在するので、１RのHTMLから、レース数を抽出する
    
    'このURLを動的に自動で取得する必要あり
    myURL = "http://race.netkeiba.com/?pid=yoso_cp&id=c201710010201"
   objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
        Sleep (100)
        StopTime = GetTickCount
        If StopTime - StartTime > 5000 Then
            str = getHTMLString(objIE)
            If InStr(str, "</html>") > 0 Then
                Exit Do
            End If
        End If
    Loop
    document_completed_flag = False
    
    Sleep (100)

    Dim i As Integer
    Dim sel As Object
    
   objIE.Visible = True    'IE を表示
    
    '0が、1・・・０～４
    '上昇度
    objIE.Document.getElementsByName("rising")(3).Click
    '逃げ・先行
    objIE.Document.getElementsByName("precede")(0).Click
    '差し・追込
    objIE.Document.getElementsByName("spurt")(0).Click
    '騎手
    objIE.Document.getElementsByName("jockey")(0).Click
    '調教師
    objIE.Document.getElementsByName("trainer")(0).Click
    '血統
    objIE.Document.getElementsByName("pedigree")(0).Click
    
    
    '送信(submit)をクリック
    For Each objTag In objIE.Document.getElementsByTagName("input")

        If InStr(objTag.outerHTML, "設定") > 0 Then

            '送信ボタンクリック
            objTag.Click

            Do While objIE.Busy = True Or objIE.ReadyState <> 4
                DoEvents
            Loop

            'ループ脱出
            Exit For
              
        End If
    Next
    
    str = getHTMLString(objIE)
    
    
    'ユニークなファイル名で保存。解析は、あとからする
    fn2 = FreeFile
    Open "c:\0213.html" For Output As #fn2
    Print #fn2, str
    
    Close #fn2
End Sub

Private Sub Command21_Click()
    Dim fn2  As Long
    Dim StartTime  As Long
    Dim StopTime  As Long
   Dim a As Object
   Dim b As Object
   Dim d As Object
   Dim c As String
   
    Set a = CreateObject("Shell.Application")
    Set b = a.Windows()
   
   'ログインしているブラウザを操作する(jun@buhi-buhi.com 名前)
    For Each d In b
        c = ""
        c = d.LocationURL
        If c = "http://www.netkeiba.com/?acc_param=top" Then
            Set objIE = d
            Exit For
        End If
    Next
   
    Sleep (100)
    DoEvents
    
    Dim str As String
    
    'http://nar.netkeiba.com/?pid=schedule&year=2015&month=8
    'スケジュールから、全URLを取得する
    'リンクは、本日以降しかないので、自力で作成する
    'http://nar.netkeiba.com/?pid=race&id=p201747022201&mode=top
    'スケジュールの印から、レースの存在を確認する
    '１Rは、確実に存在するので、１RのHTMLから、レース数を抽出する
    
    'このURLを動的に自動で取得する必要あり
    myURL = txtUrl.Text
   objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
        Sleep (100)
        StopTime = GetTickCount
        If StopTime - StartTime > 5000 Then
            str = getHTMLString(objIE)
            If InStr(str, "</html>") > 0 Then
                Exit Do
            End If
        End If
    Loop
    document_completed_flag = False
    
    Sleep (100)

End Sub


Private Sub Command22_Click()
    Dim StartTime  As Long
    Dim StopTime  As Long
   
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   objIE.Visible = True    'IE を表示
   
    Dim str As String

    myURL = "http://www.nankankeiba.com/result/1998040121010101.do"
'    myURL = "http://keiba.rakuten.co.jp/odds/sanrentan/RACEID/201702101914110501#headline"
   objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        If document_completed_flag = True Then
            Exit Do
        End If
        Sleep (100)
        StopTime = GetTickCount
    Loop
    document_completed_flag = False
    
'    Exit Sub
    
    Sleep (100)

    str = getHTMLString(objIE)

    Dim fn2  As Long
    fn2 = FreeFile
    Open "c:\0210.html" For Output As #fn2
    Print #fn2, str
    
    Close #fn2

End Sub

Private Sub Command23_Click()
    Dim aStart As String
    Dim fn2  As Long
    
    aStart = "2017/11/02 00:00:00"
'    aStart = "2016/12/21 00:00:00"
    
    'https://keiba.rakuten.co.jp/race_dividend/list/RACEID/201104220000000000
    
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   objIE.Visible = True    'IE を表示
   
    Dim str As String
    Dim aDay As String
    
    '１日全体のHTML
    '20110422 - 20170731 2012/08/20
    Do
        aDay = Format$(aStart, "yyyymmdd")
        myURL = "https://keiba.rakuten.co.jp/race_dividend/list/RACEID/" & aDay & "0000000000"
        '少佐「イシカワ、HTML取得せよっ！」
        objIE.Navigate2 myURL
        StartTime = GetTickCount
        Do While objIE.Busy = True Or objIE.ReadyState <> 4
            DoEvents
            Sleep (100)
            StopTime = GetTickCount
        Loop
        
        Sleep (100)
    
        str = getHTMLString(objIE)
        
        'イシカワ「電脳に保存しました」
        fn2 = FreeFile
        Open "c:\test\" & aDay & ".dat" For Output As #fn2
        Print #fn2, str
        
        Close #fn2
        
        '荒巻「日付を１日進めるのだ」
        aStart = DateAdd("d", 1, aStart)
        
        'タチコマ「今日だったら、ループ抜けるよっ！」
        If aDay = Format$(aStart, "yyyymmdd") = Format$(Now, "yyyymmdd") Then
            Exit Do
        End If
        
    Loop
    
    
    Exit Sub
    
    '<li class="track.+</a></li>
    '南関競馬のみ対象
    
    Do
        '少佐「イシカワ、HTML取得せよっ！」
        
        'イシカワ「電脳に保存しました」
        
        'タチコマ「今日だったら、ループ抜けるよっ！」
        
        '荒巻「日付を１日進めるのだ」
        aStart = DateAdd("d", 1, aStart)
        
    Loop
    
End Sub

Private Sub Command24_Click()
    Dim aaa As String
    Dim cnt As Long
    Dim ii As Long
    Dim dd() As String
    
    aaa = Text1.Text
    
    Call TextCodeChg(aaa)
    
    fn = FreeFile
    Open aaa For Input As #fn
    
'    Do Until EOF(fn)
        Line Input #fn, wk
'        cnt = cnt + 1
'    Loop
    
    Close #fn
    
    dd = Split(wk, vbLf)
'    ReDim dd(cnt - 1)
    
'    fn = FreeFile
'    Open aaa For Input As #fn
'
'    Do Until EOF(fn)
'        Line Input #fn, wk
'        dd(cnt) = wk
'        cnt = cnt + 1
'    Loop
'
'    Close #fn
    
    fn = FreeFile
    Open aaa & ".txt" For Output As #fn
    
    For ii = 0 To UBound(dd)
        Print #fn, dd(ii)
    Next ii
    
    Close #fn
    
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call old_chuuou
    Command25.Enabled = True
    
    MsgBox "finished!"
End Sub

Private Sub Command27_Click()
    Dim fnum As Long
    Dim wk As String
    Dim str As String
    
    fnum = FreeFile()
    
    Open "c:\test\2008092800604.htm" For Input As #fnum
    
    Do Until EOF(fnum)
        Line Input #fnum, wk
        str = str & vbLf & wk
    Loop
    
    Close #fnum
    
    Dim arr() As String
    Dim dat(1) As String
    Dim ii As Long
    
    arr = Split(str, vbLf)
    
    '<tr><td>複勝</td><td>1</td><td>120円</td><td>2</td></tr>
    For ii = 0 To UBound(arr)
        wk = InStr(arr(ii), "複勝")
        If wk > 0 Then
            dat(0) = Mid$(arr(ii), wk + 11)
            wk = InStr(dat(0), "</td>")
            dat(1) = Mid$(dat(0), wk + 9)
            dat(0) = Left$(dat(0), wk - 1)
            wk = InStr(dat(1), "</td>")
            dat(1) = Left$(dat(1), wk - 2)
            Debug.Print dat(0)
            Debug.Print dat(1)
        End If
    Next ii
    
End Sub

Private Sub Command28_Click()
    Command28.Enabled = False
    
    Dim fnum As Long
    Dim cnt As Long
    Dim wk As String
    Dim str As String
    Dim files() As String
    Dim arRace() As String
    Dim arHarai() As String
    Dim raceExist As Boolean
    Dim haraiExist As Boolean
    Dim aKaiji As String
    Dim aNiti As String
    
    'データベース確認(RACE)
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "JyoCD <='10' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "Year, MonthDay, JyoCD, RaceNum"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    cnt = -1

    Do
        If Rs.EOF = False Then
            raceExist = True
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCD = Format$(Rs("JyoCD"), "000")
            aRaceNum = Rs("RaceNum")

            cnt = cnt + 1
            ReDim Preserve arRace(cnt)
            arRace(cnt) = aYear & aMonthday & aJyoCD & aRaceNum
        Else
            Exit Do
        End If

        Rs.MoveNext

    Loop

    Rs.Close

    'データベース確認(HRAI)
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "HARAI "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "JyoCD <='10' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "Year, MonthDay, JyoCD, RaceNum"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    cnt = -1

    Do
        If Rs.EOF = False Then
            haraiExist = True
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCD = Format$(Rs("JyoCD"), "000")
            aRaceNum = Rs("RaceNum")

            cnt = cnt + 1
            ReDim Preserve arHarai(cnt)
            arHarai(cnt) = aYear & aMonthday & aJyoCD & aRaceNum
        Else
            Exit Do
        End If

        Rs.MoveNext

    Loop

    Rs.Close
    
    Dim fDat As String
    
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\res\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    files = Split(stPrompt, ",")
    
    Dim jj As Long
    Dim fukuCnt As Long
    Dim raceCnt As Long
    Dim HaraiCnt As Long
    Dim farr() As String
    Dim beforeRace As String
    
    raceCnt = 0
    HaraiCnt = 0
    
    For jj = 0 To UBound(files) - 1
        'ファイル名から情報取得(yyyymmdd jyo racenum)
        farr = Split(files(jj), "\")
        
        aYear = Left$(farr(5), 4)
        aMonthday = Mid$(farr(5), 5, 4)
        aJyoCD = Mid$(farr(5), 9, 3)
        aRaceNum = Mid$(farr(5), 12, 2)
        
        fDat = aYear & aMonthday & aJyoCD & aRaceNum
        
        fnum = FreeFile()
        
        Open files(jj) For Input As #fnum
        str = ""
        Do Until EOF(fnum)
            Line Input #fnum, wk
            str = str & vbLf & wk
        Loop
        
        Close #fnum
        
        Dim arr() As String
        Dim dat(1) As String
        Dim ii As Long
        Dim exeFlg As Boolean
        
        arr = Split(str, vbLf)
        
        fukuCnt = 0
        
        '<tr><td>複勝</td><td>1</td><td>120円</td><td>2</td></tr>
        For ii = 0 To UBound(arr)
            '回
            wk = InStr(arr(ii), "日目")
            If wk > 0 Then
                wk = InStr(arr(ii), "回")
                If wk > 0 Then
                    aKaiji = Format$(CInt(Mid$(arr(ii), wk - 2, 2)), "00")
                    wk = InStr(arr(ii), "日目")
                    aNiti = Format$(CInt(Mid$(arr(ii), wk - 1, 1)), "00")
                End If
            End If
            
            '単勝
            
            '複勝
            wk = InStr(arr(ii), "複勝")
            If wk > 0 Then
                dat(0) = Mid$(arr(ii), wk + 11)
                wk = InStr(dat(0), "</td>")
                dat(1) = Mid$(dat(0), wk + 9)
                dat(0) = Format$(Left$(dat(0), wk - 1), "00")
                wk = InStr(dat(1), "</td>")
                dat(1) = Left$(dat(1), wk - 2)
                If InStr(dat(1), ",") > 0 Then
                    dat(1) = Replace(dat(1), ",", "")
                End If
                
'                Debug.Print dat(0)
'                Debug.Print dat(1)
                fukuCnt = fukuCnt + 1
                
                
                'データベースに出力(RACE)
                'データ存在してれば、更新、なければ、追加
                exeFlg = False
                If beforeRace = fDat Then
                    exeFlg = True
                Else
                    If raceExist = True Then
                        If fDat = arRace(raceCnt) Then
                            '更新 とくにしない
                            If UBound(arRace) > raceCnt Then
                                raceCnt = raceCnt + 1
                            End If
                            exeFlg = True
                            beforeRace = fDat
                        End If
                    End If
                End If
'                If exeFlg = False Then
'                    '追加
'
'                    gstrSql = ""
'                    gstrSql = gstrSql + "SELECT "
'                    gstrSql = gstrSql + "* "
'                    gstrSql = gstrSql + "FROM "
'                    gstrSql = gstrSql + "race "
'                    gstrSql = gstrSql + "where "
'                    gstrSql = gstrSql + "year='" & aYear & "' and "
'                    gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
'                    gstrSql = gstrSql + "JyoCD='" & Right$(aJyoCD, 2) & "' and "
'                    gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
'                    ' テーブル名を指定してレコードセットを作成する
'                    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
'
'                    If Rs.EOF = True Then
'                        gstrSql = ""
'                        gstrSql = gstrSql + "insert into race (Year, monthday, jyocd, racenum"
'                        gstrSql = gstrSql + ") values ("
'
'                        gstrSql = gstrSql + "'" & aYear & "', "
'                        gstrSql = gstrSql + "'" & aMonthday & "', "
'                        gstrSql = gstrSql + "'" & Right$(aJyoCD, 2) & "', "
'                        gstrSql = gstrSql + "'" & aRaceNum & "')"
'                        '
'                        db.Execute gstrSql, dbFailOnError
'                    Else
                        gstrSql = ""
                        gstrSql = gstrSql + "update race set "
                        gstrSql = gstrSql + "Kaiji='" & aKaiji & "' "
                        gstrSql = gstrSql + "where "
                        gstrSql = gstrSql + "year='" & aYear & "' and "
                        gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
                        gstrSql = gstrSql + "JyoCD='" & Right$(aJyoCD, 2) & "' and "
                        gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
                        '
                        db.Execute gstrSql, dbFailOnError
'                    End If
'
'                    beforeRace = fDat
'                End If
                
                'データベースに出力(HARAI)
                'データベース確認(HRAI)
                gstrSql = ""
                gstrSql = gstrSql + "SELECT "
                gstrSql = gstrSql + "* "
                gstrSql = gstrSql + "FROM "
                gstrSql = gstrSql + "HARAI "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "year='" & aYear & "' and "
                gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
                gstrSql = gstrSql + "JyoCD='" & Right$(aJyoCD, 2) & "' and "
                gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
                ' テーブル名を指定してレコードセットを作成する
                Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
            
                If Rs.EOF = False Then
                    '更新 aYear & aMonthday & aJyoCD & aRaceNum
                    gstrSql = ""
                    gstrSql = gstrSql + "update harai set "
                    gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & "='" & dat(0) & "', "
                    gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & "='" & dat(1) & "' "
                    gstrSql = gstrSql + "where "
                    gstrSql = gstrSql + "year='" & aYear & "' and "
                    gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
                    gstrSql = gstrSql + "JyoCD='" & Right$(aJyoCD, 2) & "' and "
                    gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
                    '
                    db.Execute gstrSql, dbFailOnError
                Else
                    '追加
                    gstrSql = ""
                    gstrSql = gstrSql + "insert into harai (Year, monthday, jyocd, racenum, "
                    gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & ", "
                    gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & " "
                    gstrSql = gstrSql + ") values ("
                    
                    gstrSql = gstrSql + "'" & aYear & "', "
                    gstrSql = gstrSql + "'" & aMonthday & "', "
                    gstrSql = gstrSql + "'" & Right$(aJyoCD, 2) & "', "
                    gstrSql = gstrSql + "'" & aRaceNum & "', "
                    gstrSql = gstrSql + "'" & dat(0) & "', "
                    gstrSql = gstrSql + "'" & dat(1) & "')"
                    
                    db.Execute gstrSql, dbFailOnError
                End If
            
                Rs.Close
'                'データ存在してれば、更新、なければ、追加
'                exeFlg = False
'                If beforeRace = fDat Then
'                    '更新 aYear & aMonthday & aJyoCD & aRaceNum
'                    gstrSql = ""
'                    gstrSql = gstrSql + "update harai set "
'                    gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & "='" & dat(0) & "', "
'                    gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & "='" & dat(1) & "' "
'                    gstrSql = gstrSql + "where "
'                    gstrSql = gstrSql + "year='" & aYear & "' and "
'                    gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
'                    gstrSql = gstrSql + "JyoCD='" & Right$(aJyoCD, 2) & "' and "
'                    gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
'                    '
'                    db.Execute gstrSql, dbFailOnError
'
'                    If haraiExist = True Then
'                        If UBound(arHarai) > HaraiCnt Then
'                            HaraiCnt = HaraiCnt + 1
'                        End If
'                    End If
'                    exeFlg = True
'                End If
'                If exeFlg = False Then
'                    '追加
'                    gstrSql = ""
'                    gstrSql = gstrSql + "insert into harai (Year, monthday, jyocd, racenum, "
'                    gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & ", "
'                    gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & " "
'                    gstrSql = gstrSql + ") values ("
'
'                    gstrSql = gstrSql + "'" & aYear & "', "
'                    gstrSql = gstrSql + "'" & aMonthday & "', "
'                    gstrSql = gstrSql + "'" & Right$(aJyoCD, 2) & "', "
'                    gstrSql = gstrSql + "'" & aRaceNum & "', "
'                    gstrSql = gstrSql + "'" & dat(0) & "', "
'                    gstrSql = gstrSql + "'" & dat(1) & "')"
'
'                    db.Execute gstrSql, dbFailOnError
'                End If
            End If
            
            '馬連
            
            '馬単
            
            'ワイド
            
            '３連複
            
            '３連単
            
            
        Next ii
    Next jj

    Command28.Enabled = True
End Sub

Private Function cnctDB() As Long
    On Error GoTo err_handler
    
    Dim lstrDb              As String
    Dim llngRet             As Long
    
    gDB = PATH_DB
    llngRet = gfConnectDB(gDB)
    If llngRet <> 0 Then
        MsgBox "cnctDB エラー:" & llngRet
        Exit Function
    End If
    
    cnctDB = llngRet
    
    Exit Function

err_handler:
        MsgBox "cnctDB エラー:" & Err.Description & vbCr & vbLf & "エラー番号:" & Err.Number
End Function

Public Function gfConnectDB(pstrDb As String) As Long

' DAOのオブジェクト変数を宣言する
    
    ' デフォルトのワークスペースを定義する
    Set ws = DBEngine.Workspaces(0)
    ' データベースを開く
'    Set db = ws.OpenDatabase(pstrDb)
    Set db = ws.OpenDatabase(pstrDb, False, False, ";pwd=okutotta")

End Function

Private Sub Command29_Click()
    Command29.Enabled = False
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\cmpiSel\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    files = Split(stPrompt, ",")
    
    Dim jj As Long
    Dim kk As Long
    Dim umaban As String
    Dim aCmpiNinki As String
    Dim aCmpiValue As String
    Dim aR() As String
    
    raceCnt = 0
    HaraiCnt = 0
    
    For jj = 0 To UBound(files) - 1
        
        
        fDat = aYear & aMonthday & aJyoCD & aRaceNum
        
        'コンピファイルを読み込み
        fnum = FreeFile()
        
        Open files(jj) For Input As #fnum
        
        Do Until EOF(fnum)
            Line Input #fnum, wk
            aR = Split(wk, ",")
            '200701060801,3,66,12,44,16,40,10,47,8,51,5,58,13,43,6,53,4,60,1,78,7,52,2,70,14,42,15,41,11,45,9,49,,,,
            aYear = Left$(aR(0), 4)
            aMonthday = Mid$(aR(0), 5, 4)
            aJyoCD = Mid$(aR(0), 9, 2)
            aRaceNum = Mid$(aR(0), 11, 2)
            
            '人気、コンピ指数のセットで並ぶ。馬番順
            For kk = 1 To (UBound(aR) / 2) - 1
                If aR(kk) = "" Then
                    Exit For
                End If
                
                umaban = Format$(kk, "00")
                
                gstrSql = ""
                gstrSql = gstrSql + "SELECT "
                gstrSql = gstrSql + "* "
                gstrSql = gstrSql + "FROM "
                gstrSql = gstrSql + "uma_RACE "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "Year ='" & aYear & "' and "
                gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
                gstrSql = gstrSql + "JyoCD ='" & aJyoCD & "' and "
                gstrSql = gstrSql + "racenum ='" & aRaceNum & "' and "
                gstrSql = gstrSql + "umaban ='" & umaban & "'"
                
                Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
                
                If Rs.EOF = True Then
                    'データベースに追加
                    aCmpiNinki = Format$(aR(kk * 2 - 1), "00")
                    aCmpiValue = Format$(aR(kk * 2), "00")
                    gstrSql = ""
                    gstrSql = gstrSql + "insert into uma_race (Year, monthday, jyocd, racenum, umaban, CmpiNinki, CmpiValue"
                    gstrSql = gstrSql + ") values ("
                    
                    gstrSql = gstrSql + "'" & aYear & "', "
                    gstrSql = gstrSql + "'" & aMonthday & "', "
                    gstrSql = gstrSql + "'" & aJyoCD & "', "
                    gstrSql = gstrSql + "'" & aRaceNum & "', "
                    gstrSql = gstrSql + "'" & umaban & "', "
                    gstrSql = gstrSql + "'" & aCmpiNinki & "', "
                    gstrSql = gstrSql + "'" & aCmpiValue & "')"
                    
                    db.Execute gstrSql, dbFailOnError
                Else
                    gstrSql = ""
                    gstrSql = gstrSql + "update uma_race set "
                    gstrSql = gstrSql + "CmpiNinki='" & aCmpiNinki & "', "
                    gstrSql = gstrSql + "CmpiValue='" & aCmpiValue & "' "
                    gstrSql = gstrSql + "where "
                    gstrSql = gstrSql + "year='" & aYear & "' and "
                    gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
                    gstrSql = gstrSql + "JyoCD='" & aJyoCD & "' and "
                    gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' and "
                    gstrSql = gstrSql + "umaban='" & umaban & "' "
                    '
                    db.Execute gstrSql, dbFailOnError
                End If
                
                Rs.Close
                
            Next kk
            
        Loop
        
        Close #fnum
    
    Next jj
    
    Command29.Enabled = True
End Sub

Private Sub Command30_Click()
    myURL = "https://keiba.yahoo.co.jp/"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
    If chkD.Value = 1 And chkDL.Value = 0 Then
        ie.Visible = True    'IE を表示
'        ie.Visible = False
    End If
    
    Me.Caption = "Login start"
    Me.Refresh
    
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Me.Caption = "Login comp"
    Me.Refresh

End Sub

Private Sub Command31_Click()
    On Error GoTo err_hdr
    
    'https://keiba.yahoo.co.jp/race/result/1808020402/
    '                                      yyjjmmddrr
    
    Dim aDate As String
    Dim aExist As Boolean
    Dim cc As Long
    Dim cc2 As Long
    Dim cc3 As Long
    Dim str As String
    Dim iM As Integer
    Dim iY As Integer
    Dim iR As Integer
    Dim iJ As Integer
    Dim aWk As String
    Dim aPos As Long
    Dim cnt As Long
    Dim aLst() As String
    Dim aRLst() As String
    
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\yahooRes3\")
'    Set cFolder = cFso.GetFolder(App.Path & "\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    files = Split(stPrompt, ",")
    
    
restart:
    
'For iY = 2014 To 2014
'    For iM = 5 To 12
 
 myURL = "https://keiba.yahoo.co.jp/"
'起動中のIEを閉じる場合
If Not ie Is Nothing Then
   ie.Quit
   Set ie = Nothing
End If
Set ie = New SHDocVw.InternetExplorer
'指定のURLを表示
ie.Navigate2 myURL
 If chkD.Value = 1 And chkDL.Value = 0 Then
     ie.Visible = True    'IE を表示
'        ie.Visible = False
 End If
 
 Me.Caption = "Login start"
 Me.Refresh
 
 Do While ie.Busy = True Or ie.ReadyState <> 4
     DoEvents
 Loop

 Me.Caption = "Login comp"
 Me.Refresh
    
    
    'year loop
    For iY = 2018 To 2018
        'month loop
        For iM = 7 To 7
            myURL = "https://keiba.yahoo.co.jp/schedule/list/" & CStr(iY) & "/?month=" & CStr(iM)
            
            ie.Navigate2 myURL
            
            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
                Sleep 100
            Loop
            '月のURLを取得する
            str = getHTMLString(ie)
        
'            fnum2 = FreeFile()
'            ff = App.Path & "\" & "test.htm"
'            Open ff For Output As #fnum2
'
'            Print #fnum2, str
'            Close #fnum2
            
            '各開催場所のリンク情報を取得 <a href="/race/list/99060101/">(https://keiba.yahoo.co.jp/race/list/99060304/)
            cnt = -1
            Do
                aPos = InStr(str, "<a href=""/race/list/")
                If aPos > 0 Then
                    cnt = cnt + 1
                    ReDim Preserve aLst(cnt)
                    aLst(cnt) = Mid$(str, aPos + 9, 20)     '/race/list/99060101/
                    str = Mid$(str, aPos + 31)
                Else
                    Exit Do
                End If
                DoEvents
            Loop
            
            '開催場所リンク loop
            For iJ = 0 To UBound(aLst)
                
                myURL = "https://keiba.yahoo.co.jp" & aLst(iJ)
                ie.Navigate2 myURL
                Do While ie.Busy = True Or ie.ReadyState <> 4
                    If ie.Busy = False Then
                        str = getHTMLString(ie)
                        If InStr(str, "レース名") > 0 Then
                            Exit Do
                        End If
                    End If
                    DoEvents
                    Sleep 100
                Loop
                'URLを取得する
                str = getHTMLString(ie)
                
                cnt = -1
                Do
                    aPos = InStr(str, "<a href=""/race/result/")
                    If aPos > 0 Then
                        cnt = cnt + 1
                        ReDim Preserve aRLst(cnt)
                        aRLst(cnt) = Mid$(str, aPos + 9, 24)     '
                        str = Mid$(str, aPos + 31)
                    Else
                        Exit Do
                    End If
                    Sleep 100
                    DoEvents
                    
                Loop
                
                '<a href="/race/result/0005030501/">(https://keiba.yahoo.co.jp/race/result/0005030501/)
                
                'レースリンク loop
                For iR = 0 To UBound(aRLst)
                    'exist check
                    aDate = Left$(Right$(aRLst(iR), 11), 10)
                    aExist = False
                    For ii = 0 To UBound(files)
                        If aDate = Left$(Right$(files(ii), 14), 10) Then
                            aExist = True
                            Exit For
                        End If
                        
                    Next ii
                    
                    If aExist = False Then
                        myURL = "https://keiba.yahoo.co.jp" & aRLst(iR)
                        ie.Navigate2 myURL
                        Do While ie.Busy = True Or ie.ReadyState <> 4
'                            If ie.Busy = False Then
'                                str = getHTMLString(ie)
'                                If InStr(str, "単勝") > 0 Then
'                                    Exit Do
'                                End If
'                            End If
                            Sleep 100
                            DoEvents
                        Loop
                        
                        
                        'URLを取得する
                        str = getHTMLString(ie)
                        
                        
                        
                        '結果情報を保存
                        fnum2 = FreeFile()
                        ff = App.Path & "\yahooRes2\" & Mid$(aRLst(iR), 14, 10) & ".htm"
                        Open ff For Output As #fnum2
                        
                        Print #fnum2, str
                        Close #fnum2
                        
                        Debug.Print aRLst(iR)
                    End If
                    '結果情報を解析
                    '<th class="txC" rowspan="3">複勝</th>
                    '<td class="txC resultNo">3</td>
                    '<td>180円</td>
                    '<td class="resultNinki noBdrR">
                    '  <span>2番人気</span>
                    '</td>
                    '</tr>
                    '<tr>
                    '<td class="txC resultNo">4</td>
                    '<td>980円</td>
                    '<td class="resultNinki noBdrR">
                    '  <span>9番人気</span>
                    '</td>
                    '</tr>
                    '<tr>
                    '<td class="txC resultNo">10</td>
                    '<td>290円</td>
                    '<td class="resultNinki noBdrR">
                    '  <span>5番人気</span>
                    '</td>
                    '</tr>
                    '
                    '<tr>
                    '<th class="txC" rowspan="1">枠連</th>
                    
                Next iR
                
            Next iJ
            
            
        Next iM
    Next iY
    
    MsgBox "end"
    
    Exit Sub

err_hdr:
    
    Debug.Print Err.Description
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub Command32_Click()
    On Error GoTo err_hdr
    
    Dim str As String
    Dim iM As Integer
    Dim iY As Integer
    Dim iR As Integer
    Dim iJ As Integer
    Dim aWk As String
    Dim aPos As Long
    Dim cnt As Long
    Dim aLst() As String
    Dim aRLst() As String
    
    fnum2 = FreeFile()
    ff = App.Path & "\raceList.txt"
    Open ff For Append As #fnum2
    
    'year loop
    For iY = 2013 To 2017
        'month loop
        For iM = 1 To 12
            
            myURL = "https://keiba.yahoo.co.jp/schedule/list/" & CStr(iY) & "/?month=" & CStr(iM)
            
            ie.Navigate2 myURL
            
            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop
            '月のURLを取得する
            str = getHTMLString(ie)
        
            '各開催場所のリンク情報を取得 <a href="/race/list/99060101/">(https://keiba.yahoo.co.jp/race/list/99060304/)
            cnt = -1
            Do
                aPos = InStr(str, "<a href=""/race/list/")
                If aPos > 0 Then
                    cnt = cnt + 1
                    ReDim Preserve aLst(cnt)
                    aLst(cnt) = Mid$(str, aPos + 9, 20)     '/race/list/99060101/
                    str = Mid$(str, aPos + 31)
                Else
                    Exit Do
                End If
                
            Loop
            
            '開催場所リンク loop
            For iJ = 0 To UBound(aLst)
                
                myURL = "https://keiba.yahoo.co.jp" & aLst(iJ)
                ie.Navigate2 myURL
                Do While ie.Busy = True Or ie.ReadyState <> 4
                    DoEvents
                Loop
                'URLを取得する
                str = getHTMLString(ie)
                
                cnt = -1
                Do
                    aPos = InStr(str, "<a href=""/race/result/")
                    If aPos > 0 Then
                        cnt = cnt + 1
                        ReDim Preserve aRLst(cnt)
                        aRLst(cnt) = Mid$(str, aPos + 9, 24)     '
                        str = Mid$(str, aPos + 31)
                    Else
                        Exit Do
                    End If
                    
                Loop
                
                '<a href="/race/result/0005030501/">(https://keiba.yahoo.co.jp/race/result/0005030501/)
                
                'レースリンク loop
                For iR = 0 To UBound(aRLst)
                    '結果情報を保存
                    
                    Print #fnum2, aRLst(iR)
                    
                    Debug.Print aRLst(iR)
                Next iR
                
            Next iJ
            
        Next iM
    Next iY
    
    Close #fnum2
    
    MsgBox "finish!"
    
    Exit Sub

err_hdr:
    Close #fnum2
    
    Debug.Print Err.Description
    MsgBox Err.Description
    Exit Sub

End Sub


Private Sub Command33_Click()
    Dim aArr() As String
    Dim aCnt As Long
    Dim wk As String
    Dim aY As String
    Dim aM As String
    Dim aJ As String
    Dim aR As String
    Dim aK As String
    Dim aN As String
    Dim aBfr As String
    
    aCnt = -1
    
    fnum2 = FreeFile()
    ff = App.Path & "\raceList.txt"
    Open ff For Input As #fnum2
    
    Do Until EOF(fnum2)
        Line Input #fnum2, wk
        aCnt = aCnt + 1
    Loop
    
    Close #fnum2
    
    ReDim aArr(aCnt)
    aCnt = -1
    
    fnum2 = FreeFile()
    ff = App.Path & "\raceList.txt"
    Open ff For Input As #fnum2
    
    Do Until EOF(fnum2)
        Line Input #fnum2, wk
        aCnt = aCnt + 1
        aArr(aCnt) = wk
    Loop
    
    Close #fnum2
    
    fnum2 = FreeFile()
    ff = App.Path & "\check.txt"
    Open ff For Append As #fnum2
    
    aBfr = ""
    
    For ii = 0 To aCnt
        '/race/result/0706010101/
        '/race/result/YYJJKKNNRR/
        wk = Mid$(aArr(ii), 14, 10)
        aY = "20" & Left$(wk, 2)
        aJ = Mid$(wk, 3, 2)
        aK = Mid$(wk, 5, 2)
        aN = Mid$(wk, 5, 2)
        aR = Right$(wk, 2)
        
        'RACE check
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "RACE "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "year ='" & aY & "' and "
        gstrSql = gstrSql + "monthday ='" & aM & "' and "
        gstrSql = gstrSql + "JyoCD ='" & aJ & "' and "
        gstrSql = gstrSql + "racenum ='" & aR & "'"
        
        ' テーブル名を指定してレコードセットを作成する
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
        If Rs.EOF = False Then
        Else
            Print #fnum2, "race:" & wk
        End If
    
        Rs.Close
        
        'HARAI check
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "harai "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "year ='" & aY & "' and "
        gstrSql = gstrSql + "monthday ='" & aM & "' and "
        gstrSql = gstrSql + "JyoCD ='" & aJ & "' and "
        gstrSql = gstrSql + "racenum ='" & aR & "'"
        
        ' テーブル名を指定してレコードセットを作成する
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
        If Rs.EOF = False Then
        Else
            Print #fnum2, "harai:" & wk
        End If
    
        Rs.Close
        
        'UMA_RACE check
        gstrSql = ""
        gstrSql = gstrSql + "SELECT "
        gstrSql = gstrSql + "* "
        gstrSql = gstrSql + "FROM "
        gstrSql = gstrSql + "UMA_RACE "
        gstrSql = gstrSql + "where "
        gstrSql = gstrSql + "year ='" & aY & "' and "
        gstrSql = gstrSql + "monthday ='" & aM & "' and "
        gstrSql = gstrSql + "JyoCD ='" & aJ & "' and "
        gstrSql = gstrSql + "racenum ='" & aR & "'"
        
        ' テーブル名を指定してレコードセットを作成する
        Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
        If Rs.EOF = False Then
        Else
            Print #fnum2, "harai:" & wk
        End If
    
        Rs.Close
    Next ii
    
    Close #fnum2
    
End Sub

Private Sub Command34_Click()
    Dim aArr() As String
    Dim aCnt As Long
    Dim wk As String
    Dim aTrack As String
    Dim aJyokenCD5 As String
    Dim aY As String
    Dim aM As String
    Dim aD As String
    Dim aJ As String
    Dim aR As String
    Dim aK As String
    Dim aN As String
    Dim aBfr As String
    Dim aPY As String
    Dim aPM As String
    Dim aPD As String
    Dim aUma As String
    Dim aHarai As String
    Dim aWk1 As Long
    Dim aWk2 As Long
    Dim aWk3 As Long
    Dim aZan As String
    Dim aNext As Long
    Dim fpos As Long
    
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\yahooRes3\")
'    Set cFolder = cFso.GetFolder(App.Path & "\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    files = Split(stPrompt, ",")
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
        
    For ii = 0 To UBound(files) - 1   'yyjjkknnrr.htm
        aY = "20" & Left$(Right$(files(ii), 14), 2)
        aJ = Mid$(Right$(files(ii), 14), 3, 2)
        aR = Mid$(Right$(files(ii), 14), 9, 2)
        aCnt = -1
        
        fnum2 = FreeFile()
        ff = files(ii)
        Open ff For Input As #fnum2
        
        Line Input #fnum2, wk
        
        wk = Replace(wk, vbLf, "")
        
        Close #fnum2
        
        '正規表現オブジェクトの宣言
        Set objRegExp = New RegExp
        
        With objRegExp
            .Global = True '複数マッチ可
            .IgnoreCase = True
            .Global = True
            .MultiLine = True
            
            'raceTitDay">2014年3月2日
            '回小倉8日
            'th class="txC ****** <th class="txC
            'resultNo">6</td><td>610円
            
            '日付
            .Pattern = "....年[0-9]+月[0-9]+日"
            
            pos = 0
            retstr = ""
            Set Matches = .Execute(wk)   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                pos = Match.FirstIndex       '一致する文字列が見つかった位置
                retstr = Match.Value
'                Debug.Print retstr
                Exit For
            Next
            
            If pos > 0 Then
                aPY = InStr(retstr, "年")
                aPM = InStr(retstr, "月")
                aPD = Format$(InStr(retstr, "日"), "00")
                aM = Format$(Mid$(retstr, aPY + 1, aPM - aPY - 1), "00")
                aD = Format$(Mid$(retstr, aPM + 1, aPD - aPM - 1), "00")
                
                If (aM & aD) <> "" Then
                '複勝
            '    .Pattern = "th class=""txC.+\<th class=""txC"
                .Pattern = "複勝.+円"
            '    .Pattern = "th class=""txC.+"
                
                fpos = 0
                retstr = ""
                Set Matches = .Execute(wk)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                    fpos = Match.FirstIndex       '一致する文字列が見つかった位置
                    retstr = Match.Value
'                    Debug.Print retstr
                    Exit For
                Next
                
                Caption = aY & aM & aJ & aR
                Me.Refresh
                DoEvents
                
                .Pattern = "raceTitMeta.+\</span\>" ' \</p\>"
'                    .Pattern = "raceTitMeta.+\</span\> \</p\>"
                aTrack = ""
                Set Matches = .Execute(wk)   ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                    pos = Match.FirstIndex       '一致する文字列が見つかった位置
                    aTrack = Match.Value
                    Exit For
                Next
                If aTrack <> "" Then
                    aTrack = aTrack
                End If
                
                If InStr(aTrack, "新馬") > 0 Then
                    aTrack = ""
                    aJyokenCD5 = "701"
                ElseIf InStr(aTrack, "障害") > 0 Then
                    aTrack = "52"
                    aJyokenCD5 = ""
                Else
                    aTrack = "10"
                    aJyokenCD5 = "000"
                End If
                
                gstrSql = ""
                gstrSql = gstrSql + "SELECT "
                gstrSql = gstrSql + "* "
                gstrSql = gstrSql + "FROM "
                gstrSql = gstrSql + "race "
                gstrSql = gstrSql + "where "
                gstrSql = gstrSql + "year='" & aY & "' and "
                gstrSql = gstrSql + "monthday='" & aM & aD & "' and "
                gstrSql = gstrSql + "JyoCD='" & aJ & "' and "
                gstrSql = gstrSql + "RaceNum='" & aR & "' "
                ' テーブル名を指定してレコードセットを作成する
                Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

                If Rs.EOF = True Then
                    gstrSql = ""
                    gstrSql = gstrSql + "insert into race (Year, monthday, jyocd, racenum, TrackCD, JyokenCD5"
                    gstrSql = gstrSql + ") values ("

                    gstrSql = gstrSql + "'" & aY & "', "
                    gstrSql = gstrSql + "'" & aM & aD & "', "
                    gstrSql = gstrSql + "'" & aJ & "', "
                    gstrSql = gstrSql + "'" & aR & "', "
                    gstrSql = gstrSql + "'" & aTrack & "', "
                    gstrSql = gstrSql + "'" & aJyokenCD5 & "')"
                    '
                    db.Execute gstrSql, dbFailOnError
                Else
                    gstrSql = ""
                    gstrSql = gstrSql + "update race set "
                    gstrSql = gstrSql + "TrackCD='" & aTrack & "', "             '傷害
                    gstrSql = gstrSql + "JyokenCD5='" & aJyokenCD5 & "' "          '新馬
                    gstrSql = gstrSql + "where "
                    gstrSql = gstrSql + "year='" & aY & "' and "
                    gstrSql = gstrSql + "monthday='" & aM & aD & "' and "
                    gstrSql = gstrSql + "JyoCD='" & aJ & "' and "
                    gstrSql = gstrSql + "RaceNum='" & aR & "' "
                    '
                    db.Execute gstrSql, dbFailOnError
                End If
                
                Rs.Close
                    
                If fpos > 0 Then
                    aZan = retstr
                    fukuCnt = 0
                    
                    Do
                        'bpm 180
                        '複勝</th><td class="txC resultNo">8</td><td>410円
                        aNext = InStr(aZan, "<th class=""txC"" rowspan")
                        aWk1 = InStr(aZan, "resultNo")
                        aWk2 = InStr(aZan, "</td><td>")
                        aWk3 = InStr(aZan, "円")
                        aUma = Format$(Mid$(aZan, aWk1 + 10, aWk2 - (aWk1 + 10)), "00")
                        aHarai = Replace(Mid$(aZan, aWk2 + 9, aWk3 - (aWk2 + 9)), ",", "")
                        
                        aZan = Mid$(aZan, aWk3 + 1)
                        
                        If aNext < aWk1 Then
                            Exit Do
                        Else
                        
                            gstrSql = ""
                            gstrSql = gstrSql + "SELECT "
                            gstrSql = gstrSql + "* "
                            gstrSql = gstrSql + "FROM "
                            gstrSql = gstrSql + "HARAI "
                            gstrSql = gstrSql + "where "
                            gstrSql = gstrSql + "year='" & aY & "' and "
                            gstrSql = gstrSql + "monthday='" & aM & aD & "' and "
                            gstrSql = gstrSql + "JyoCD='" & aJ & "' and "
                            gstrSql = gstrSql + "RaceNum='" & aR & "' "
                            ' テーブル名を指定してレコードセットを作成する
                            Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
                            
                            fukuCnt = fukuCnt + 1
                            
                            If Rs.EOF = False Then
                                '更新 aYear & aMonthday & aJyoCD & aRaceNum
                                gstrSql = ""
                                gstrSql = gstrSql + "update harai set "
                                gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & "='" & aUma & "', "
                                gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & "='" & aHarai & "' "
                                gstrSql = gstrSql + "where "
                                gstrSql = gstrSql + "year='" & aY & "' and "
                                gstrSql = gstrSql + "monthday='" & aM & aD & "' and "
                                gstrSql = gstrSql + "JyoCD='" & aJ & "' and "
                                gstrSql = gstrSql + "RaceNum='" & aR & "' "
                                '
                                db.Execute gstrSql, dbFailOnError
                            Else
                                '追加
                                gstrSql = ""
                                gstrSql = gstrSql + "insert into harai (Year, monthday, jyocd, racenum, "
                                gstrSql = gstrSql + "PayFukusyoUmaban" & CStr(fukuCnt) & ", "
                                gstrSql = gstrSql + "PayFukusyoPay" & CStr(fukuCnt) & " "
                                gstrSql = gstrSql + ") values ("
                                
                                gstrSql = gstrSql + "'" & aY & "', "
                                gstrSql = gstrSql + "'" & aM & aD & "', "
                                gstrSql = gstrSql + "'" & aJ & "', "
                                gstrSql = gstrSql + "'" & aR & "', "
                                gstrSql = gstrSql + "'" & aUma & "', "
                                gstrSql = gstrSql + "'" & aHarai & "')"
                                
                                db.Execute gstrSql, dbFailOnError
                            End If
                        
                            Rs.Close
                        End If
                    Loop
                End If
                End If
                
            End If
        
        End With
    Next ii
    
    
'id="raceTitDay">2014年3月2日（日） <span>|</span> 1回小倉8日 <span>|</span> 14:50発走</p>
'<h1 class="fntB">
'関門橋ステークス</h1>
'<p class="fntSS gryB" id="raceTitMeta">芝・右 1800m <span>|</span> 天気：<img width="15" height="15" class="spBg kumori" alt="曇" src="https://s.yimg.jp/images/clear.gif" border="0"> <span>|</span> 馬場：<img width="25" height="15" class="spBg ryou" alt="良" src="https://s.yimg.jp/images/clear.gif" border="0"> <span>|</span> サラ系4歳以上 <span>|</span> 1600万下 （混合） 定量 <span>|</span> 本賞金：1590、640、400、240、159万円 <span>|</span>
    
   
'<th class="txC" rowspan="3">
'    '複勝</th>
'<td class="txC resultNo">6</td>
'<td>610円</td>ass="resultNinki noBdrR">
'  <span>8番人気</span>
'</td>
'</tr>
'<tr>
'<td class="txC resultNo">8</td>
'<td>140円</td>
'<td class="resultNinki noBdrR">
'  <span>1番人気</span>
'</td>
'</tr>
'<tr>
'<td class="txC resultNo">9</td>
'<td>660円</td>
'<td class="resultNinki noBdrR">
'  <span>9番人気</span>
'</td>
'</tr>
'
'<tr>
'<th class="txC" rowspan="1">枠連<
    MsgBox "finish!"
End Sub



Private Sub Command35_Click()
    
    Dim fn2  As Long
    fn2 = FreeFile
    Open App.Path & "\check.txt" For Output As #fn2
    
    'race から harai,race_umaの有無をチェック
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "RACE "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "JyoCD <='10' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "Year, MonthDay, JyoCD, RaceNum"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCD = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            
            'harai
            gstrSql = ""
            gstrSql = gstrSql + "SELECT "
            gstrSql = gstrSql + "* "
            gstrSql = gstrSql + "FROM "
            gstrSql = gstrSql + "HARAI "
            gstrSql = gstrSql + "where "
            gstrSql = gstrSql + "year='" & aYear & "' and "
            gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
            gstrSql = gstrSql + "JyoCD='" & aJyoCD & "' and "
            gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
            ' テーブル名を指定してレコードセットを作成する
            Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
            Do
                If Rs2.EOF = False Then
                Else
                    Print #fn2, "harai:" & aYear & aMonthday & aJyoCD & aRaceNum
                End If
        
                Exit Do
            Loop
        
            Rs2.Close
            
            'uma_race
            gstrSql = ""
            gstrSql = gstrSql + "SELECT "
            gstrSql = gstrSql + "* "
            gstrSql = gstrSql + "FROM "
            gstrSql = gstrSql + "uma_RACE "
            gstrSql = gstrSql + "where "
            gstrSql = gstrSql + "Year ='" & aYear & "' and "
            gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
            gstrSql = gstrSql + "JyoCD ='" & aJyoCD & "' and "
            gstrSql = gstrSql + "racenum ='" & aRaceNum & "' "
            
            Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
            Do
                If Rs2.EOF = False Then
                Else
                    Print #fn2, "uma:" & aYear & aMonthday & aJyoCD & aRaceNum
                End If
        
                Exit Do
            Loop
        
            Rs2.Close
        Else
            Exit Do
        End If

        Rs.MoveNext

    Loop

    Rs.Close
    
    MsgBox "finish"
    
    Close #fn2
    
End Sub

Private Sub Command36_Click()
    
    Dim fn2  As Long
    fn2 = FreeFile
    Open App.Path & "\check2.txt" For Output As #fn2
    
    'race_uma から harai,raceの有無をチェック
    
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_RACE "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "umaban ='01' "
    gstrSql = gstrSql + "ORDER BY "
    gstrSql = gstrSql + "Year, MonthDay, JyoCD, RaceNum"
    ' テーブル名を指定してレコードセットを作成する
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)

    Do
        If Rs.EOF = False Then
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCD = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            
            'harai
            gstrSql = ""
            gstrSql = gstrSql + "SELECT "
            gstrSql = gstrSql + "* "
            gstrSql = gstrSql + "FROM "
            gstrSql = gstrSql + "HARAI "
            gstrSql = gstrSql + "where "
            gstrSql = gstrSql + "year='" & aYear & "' and "
            gstrSql = gstrSql + "monthday='" & aMonthday & "' and "
            gstrSql = gstrSql + "JyoCD='" & aJyoCD & "' and "
            gstrSql = gstrSql + "RaceNum='" & aRaceNum & "' "
            ' テーブル名を指定してレコードセットを作成する
            Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
            Do
                If Rs2.EOF = False Then
                Else
                    Print #fn2, "harai:" & aYear & aMonthday & aJyoCD & aRaceNum
                End If
        
                Exit Do
            Loop
        
            Rs2.Close
            
            'uma_race
            gstrSql = ""
            gstrSql = gstrSql + "SELECT "
            gstrSql = gstrSql + "* "
            gstrSql = gstrSql + "FROM "
            gstrSql = gstrSql + "RACE "
            gstrSql = gstrSql + "where "
            gstrSql = gstrSql + "Year ='" & aYear & "' and "
            gstrSql = gstrSql + "Monthday ='" & aMonthday & "' and "
            gstrSql = gstrSql + "JyoCD ='" & aJyoCD & "' and "
            gstrSql = gstrSql + "racenum ='" & aRaceNum & "' "
            
            Set Rs2 = db.OpenRecordset(gstrSql, dbOpenDynaset)
        
            Do
                If Rs2.EOF = False Then
                Else
                    Print #fn2, "race:" & aYear & aMonthday & aJyoCD & aRaceNum
                End If
        
                Exit Do
            Loop
        
            Rs2.Close
        Else
            Exit Do
        End If

        Rs.MoveNext

    Loop

    Rs.Close
    
    MsgBox "finish"
    
    Close #fn2

End Sub


Private Sub Command37_Click()
    gstrSql = ""
    gstrSql = gstrSql + "SELECT "
    gstrSql = gstrSql + "* "
    gstrSql = gstrSql + "FROM "
    gstrSql = gstrSql + "uma_RACE "
    gstrSql = gstrSql + "where "
    gstrSql = gstrSql + "Year ='2011' and "
    gstrSql = gstrSql + "Monthday ='0212' and "
    gstrSql = gstrSql + "JyoCD ='10' and "
    gstrSql = gstrSql + "racenum >='04' "
    
    Set Rs = db.OpenRecordset(gstrSql, dbOpenDynaset)
    
    Do
        If Rs.EOF = False Then
            aYear = Rs("year")
            aMonthday = Rs("MonthDay")
            aJyoCD = Rs("JyoCD")
            aRaceNum = Rs("RaceNum")
            aUmaban = Rs("umaban")
            aCmpiNinki = Rs("CmpiNinki")
            aCmpiValue = Rs("CmpiValue")
            
            'データベースに追加
            gstrSql = ""
            gstrSql = gstrSql + "insert into uma_race (Year, monthday, jyocd, racenum, umaban, CmpiNinki, CmpiValue"
            gstrSql = gstrSql + ") values ("
            
            gstrSql = gstrSql + "'" & aYear & "', "
            gstrSql = gstrSql + "'0213', "
            gstrSql = gstrSql + "'" & aJyoCD & "', "
            gstrSql = gstrSql + "'" & aRaceNum & "', "
            gstrSql = gstrSql + "'" & aUmaban & "', "
            gstrSql = gstrSql + "'" & aCmpiNinki & "', "
            gstrSql = gstrSql + "'" & aCmpiValue & "')"
            
            db.Execute gstrSql, dbFailOnError
        Else
            Exit Do
        End If
        
        Rs.MoveNext
    Loop
    
    Rs.Close
    
    
    MsgBox "finish"
End Sub

Private Sub Command38_Click()
    myURL = "http://p.nikkansports.com/goku-uma/member/compi/compi_db.zpl#/index/?CompiMin1=40&CompiMax1=90&CompiMin2=40&CompiMax2=90&CompiMin3=40&CompiMax3=90&Compi1=&Compi2=&Compi3=&Compi4=&Compi5=8&Compi6=&Compi7=&Compi8=51&Compi9=&Compi10=&Compi11=&Compi12=&Compi13=&Compi14=&Compi15=&Compi16=&Compi17=&Compi18=&StartYear=2007&StartMonth=1&StartDay=1&EndYear=2016&EndMonth=12&EndDay=31&DistanceMin=0&DistanceMax=3600&BettingType=1&PayoffMin=0&PayoffMax=100000000&DiffCompiRankMin=&DiffCompiRankMax=&DiffMin=&DiffMax=&HeadsMin=&HeadsMax="
   '起動中のIEを閉じる場合
'   If Not ie Is Nothing Then
'      ie.Quit
'      Set ie = Nothing
'   End If
'   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
'   ie.Visible = True    'IE を表示
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
   
   
    For Each objTag In ie.Document.getElementsByTagName("input")

        If InStr(objTag.outerHTML, "検索開始") > 0 Then

            '送信ボタンクリック
            objTag.Click

            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop

            'ループ脱出
            Exit For
              
        End If
    Next
   
   
   
'   With ie
''      .Document.Forms(0).elements("mailAddress").Value = "jun@buhi-buhi.com"
''      .Document.Forms(0).elements("password").Value = "two784"
'
''Sleep (10)
'
''      .Document.getElementsByTagName("INPUT")(2).submit
''      .Document.Forms(0).elements(2).Click
'      .Document.getElementsByTagName("input")(28).Click
'   End With

End Sub

Private Sub Command39_Click()
    Dim i As Integer
    Dim src As String
    Dim file As String
    Dim wfile As String
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
        Call analRakutenHarai(src & ".txt")
    Next i

    MsgBox "finish"
End Sub

Private Function analRakutendetailHarai(src As String)
Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
Dim fnTfr As Integer
Dim fn As Integer
Dim wk As String
Dim wk2 As String
Dim wk3 As String
Dim lCnt As Integer
Dim aNum As Integer
Dim Data() As String
Dim allDat As String
Dim aDay As String

    fn = FreeFile
    Open src For Input As #fn
    
    lCnt = 0
    Do Until EOF(fn)
        Line Input #fn, wk
        ReDim Preserve Data(lCnt)
        Data(lCnt) = wk
        lCnt = lCnt + 1
    Loop
    
    '201708012135050200
    'yyyymmddjj??kknn??
    aDay = Left$(Right$(src, 26), 18)
    
    '<<ファイル 閉>>
    Close #fn
    
    'wfn = FreeFile
    'Open wfile For Append As #wfn
    
    Open src & ".result.txt" For Output As #fn
    
    '<<データ解析>>
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim aPhase As Integer
    Dim aRace As String
    Dim aUma(5) As String
    Dim aFuku(5) As String
    Dim ii As Integer
    
    aPhase = 1  'レース検索
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    
    '<h3 class="headline"><span>■</span>1R
    '<th scope="row">複勝</th>
    '<td class="number">6<br>1</td>
    '<td class="money">360 円<br>190 円</td>
    
    For lCnt = 0 To UBound(Data)
         
         Select Case aPhase
         Case 1
            retstr = ""
            .Pattern = "<h3 class=""headline""><span>■</span>[0-9]+R"
            pos = 0
            Set Matches = .Execute(Data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.Value
            
            Next
            
            If retstr <> "" Then
                If Left$(Right$(retstr, 3), 1) = ">" Then
                    aRace = Mid$(Right$(retstr, 3), 2, 1)
                Else
                    aRace = Mid$(Right$(retstr, 3), 1, 2)
                End If
                allDat = allDat & aDay & "," & aRace & ","
                aPhase = 2
                For ii = 0 To UBound(aUma)
                    aUma(ii) = ""
                    aFuku(ii) = ""
                Next ii
            End If
            
         Case 2
            retstr = ""
            .Pattern = "<th scope=""row"">複勝</th>"
            pos = 0
            Set Matches = .Execute(Data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.Value
            
            Next
            
            If retstr <> "" Then
                aPhase = 3
            End If
         Case 3
            '<td class="number">
            retstr = ""
            .Pattern = "<td class=""number"">.+</td>"
            pos = 0
            Set Matches = .Execute(Data(lCnt))   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.Value
            
            Next
            If retstr <> "" Then
                retstr = ""
                '馬番を取得
                .Pattern = "[0-9]+<"
                pos = 0
                aNum = 0
                Set Matches = .Execute(Data(lCnt))  ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                    pos = Match.FirstIndex       '一致する文字列が見つかった位置
                    retstr = Match.Value
                    aUma(aNum) = Left$(retstr, Len(retstr) - 1)
                    allDat = allDat & aUma(aNum) & "/"
                    aNum = aNum + 1
                Next
                
                aPhase = 4
            End If
                
         Case 4
            '払戻を取得
            retstr = ""
            .Pattern = ">.+円<"
            pos = 0
            Set Matches = .Execute(Data(lCnt))  ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               retstr = Match.Value
            
            Next
                
            If retstr <> "" Then
                retstr = ""
                '払戻を取得
                .Pattern = "[0-9, ]+円"
'                .Pattern = ".+円"
                pos = 0
                allDat = allDat & ","
                aNum = 0
                Set Matches = .Execute(Data(lCnt))  ' 検索を実行します。
                For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                    pos = Match.FirstIndex       '一致する文字列が見つかった位置
                    retstr = Match.Value
                    aFuku(aNum) = Left$(retstr, Len(retstr) - 2)
                    allDat = allDat & aFuku(aNum) & "/"
                    aNum = aNum + 1
                
                Next
                
                allDat = allDat & aFuku(aNum) & vbCrLf
                aPhase = 1
            End If
         End Select
        
        
    Next lCnt
End With
    
    Print #fn, allDat
    Close #fn
    Set objRegExp = Nothing
    

End Function

'Scrrun.dll
Private Function analRakutenHarai(src As String)
Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
Dim fnTfr As Integer
Dim fn As Integer
Dim wk As String
Dim wk2 As String
Dim wk3 As String
Dim lCnt As Integer
Dim Data() As String
    
    fn = FreeFile
    Open src For Input As #fn
    
    lCnt = 0
    Do Until EOF(fn)
        Line Input #fn, wk
        ReDim Preserve Data(lCnt)
        Data(lCnt) = wk
        lCnt = lCnt + 1
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    'wfn = FreeFile
    'Open wfile For Append As #wfn
    
    
    '<<データ解析>>
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    fnTfr = FreeFile
    Open "urlList.txt" For Append As #fnTfr
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    
    '<a href="/race_dividend/list/RACEID/201104222015020500">大　井</a></li>
    'https://keiba.rakuten.co.jp/race_dividend/list/RACEID/201104222015020500
    
    For lCnt = 0 To UBound(Data)
         '１日全体の各開催場所のURLを割り出し
         .Pattern = "<a href=""/race_dividend/list/RACEID.+</a></li>"
         
        pos = 0
        Set Matches = .Execute(Data(lCnt))   ' 検索を実行します。
        For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
           pos = Match.FirstIndex       '一致する文字列が見つかった位置
           retstr = Match.Value
            Print #fnTfr, "https://keiba.rakuten.co.jp" & Mid$(retstr, 10, 45)
        
        Next
        
        
        
    Next lCnt
End With
    
    Close #fnTfr
    Set objRegExp = Nothing
    
End Function

Private Sub Command40_Click()
    Dim i As Integer
    Dim src As String
    Dim file As String
    Dim wfile As String
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
        Call analRakutendetailHarai(src & ".txt")
    Next i

    MsgBox "finish"
End Sub

Private Sub Command41_Click()
    Dim arrwk() As String
    Dim wk As String
    Dim wk2 As String
    Dim fn  As Long
    
    fn = FreeFile
    Open "urlList.txt" For Input As #fn
    
    '<<ファイル 読>>
    Line Input #fn, wk2
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wk2 = wk2 & "," & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    arrwk = Split(wk2, ",")
    
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   objIE.Visible = True    'IE を表示
   
    Dim str As String
    Dim aDay As String
    Dim ii As Integer
    
    '20110422 - 20170731 2012/08/20
    For ii = 0 To UBound(arrwk)
        myURL = arrwk(ii)
        '少佐「イシカワ、HTML取得せよっ！」
        objIE.Navigate2 myURL
        StartTime = GetTickCount
        Do While objIE.Busy = True Or objIE.ReadyState <> 4
            DoEvents
            Sleep (100)
            StopTime = GetTickCount
        Loop
        
        Sleep (100)
    
        str = getHTMLString(objIE)
        
        'イシカワ「電脳に保存しました」
        fn2 = FreeFile
        Open "c:\test2\" & Right$(arrwk(ii), 18) & ".txt" For Output As #fn2
        Print #fn2, str
        
        Close #fn2
        
    Next ii
    
    MsgBox "Finish"
    Exit Sub

End Sub

Private Sub Command42_Click()
    myURL = "https://keiba.yahoo.co.jp/schedule/list/"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
   ie.Visible = True    'IE を表示
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Dim str As String
    Dim str2 As String
    
    str = getHTMLString(ie)

'    Debug.Print str
'    Text2.Text = str
    
    Dim fn As Long
    Dim src As String
    
'    src = "c:\temp\togetu.txt"
'    fn = FreeFile
'    Open src For Output As #fn
'    Print #fn, str
'    Close #fn
    Dim aDat() As String
    Dim ii As Integer
    Dim jj As Integer
    Dim aWkStr As String
    Dim aJyoURL() As String
    Dim aJyoCnt As Integer
    Dim aUrlPos As Integer
    Dim aWkPos As Integer
    
    aDat = Split(str, vbLf)
    '当日URLをHTMLから日付を取得 -> >8日（
    aWkStr = ">" & Format$(Now, "d") & "日（"
    For ii = 0 To UBound(aDat)
        If InStr(aDat(ii), aWkStr) > 0 Then
            ReDim Preserve aJyoURL(aJyoCnt)
            aUrlPos = InStr(aDat(ii), "href")
            aJyoURL(aJyoCnt) = "https://keiba.yahoo.co.jp" & Mid$(aDat(ii), aUrlPos + 6, 20)
            aJyoCnt = aJyoCnt + 1
        End If
    Next ii
    
    Dim aRaceNum As String
    Dim aJyoCode As String
    Dim aHassoTime As String
    Dim aBasicCnt As Integer
    Dim aRaceURL As String
    Dim aJyoLink As String
    
    For ii = 0 To UBound(aJyoURL)
        aJyoCode = Mid$(aJyoURL(ii), 39, 2)
        aJyoLink = Left$(Right$(aJyoURL(ii), 9), 8)
        
        myURL = aJyoURL(ii)
        ie.Navigate2 myURL
        Do While ie.Busy = True Or ie.ReadyState <> 4
            DoEvents
        Loop
        
        str = getHTMLString(ie)
        
        aDat = Split(str, vbLf)
        
        For jj = 0 To UBound(aDat)
            aWkStr = "scheRNo"
            aWkPos = InStr(aDat(jj), aWkStr)
            If aWkPos > 0 Then
                aRaceNum = Mid$(aDat(jj), aWkPos + 9, 2)  'race No.
                If Right$(aRaceNum, 1) = "R" Then
                    aRaceNum = Left$(aRaceNum, 1)
                End If
                aWkPos = InStr(aDat(jj), "fntSS")
                aWkStr = Mid$(aDat(jj), aWkPos + 7)
                aWkPos = InStr(aWkStr, "<")
                aHassoTime = Left$(aWkStr, aWkPos - 1)       'time
                
                ReDim Preserve aBasicDat(aBasicCnt)
                aBasicDat(aBasicCnt) = aJyoCode & "," & aRaceNum & "," & aHassoTime & "," & aJyoLink
                aBasicCnt = aBasicCnt + 1
            End If
        Next jj
    Next ii
    
    Timer1.Enabled = True
    
    Exit Sub
    
Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
    Dim aResUrl() As String
    Dim cnt As Long
    Dim cnt2 As Long
    Dim wA As Integer
    
    Dim raceNum As Integer
    Dim maxRaceNum As Integer
    Dim lstUrl As String
    Dim resGet As String
    Dim resLp As Integer
    Dim resMny() As String
    Dim wkPrt As String
    Dim pt1 As Integer
    Dim pt2 As Integer
    Dim wkwk As String
    
    cnt = -1
    
'<<データ解析>>
'正規表現オブジェクトの宣言
Set objRegExp = New RegExp

With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    '各年毎の各月のデータ取得
     .Pattern = "a href=""" & "/" & "race" & "/" & "result" & "/........../"
    
    pos = 0
    Set Matches = .Execute(str)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
        ReDim Preserve aResUrl(cnt)
        aResUrl(cnt) = "http://keiba.yahoo.co.jp" & Mid$(retstr, 9)
        aResUrl(cnt) = Left$(aResUrl(cnt), Len(aResUrl(cnt)) - 3)
'       Debug.Print retstr
    Next
    
    raceNum = 1
    
    For wA = 0 To cnt
        '任意の日、開催場所のHTMLから、各レース毎のURLを取得
        
        'レースMax.を取得
        lstUrl = Left$(aResUrl(wA), 30) & "list/" & Right$(aResUrl(wA), 9)
        ie.Navigate2 lstUrl
'        Do While ie.Busy = True Or ie.ReadyState <> 4
'            Call Sleep(1)
'            DoEvents
'        Loop
        Do While ie.Busy = True Or ie.ReadyState <> 4
'            Call Sleep(1)
'            If ie.Busy = False Then
'                Exit Do
'            End If
'            If ie.ReadyState = READYSTATE_COMPLETE Then
'                Exit Do
'            End If
            
            DoEvents
        Loop
        
        'scheRNo">.+R</p>
         .Pattern = "scheRNo""\>.+R\</p\>"
        str2 = ""
        str2 = getHTMLString(ie)
'    fn = FreeFile
'    Open src For Output As #fn
'    Print #fn, str2
'    Close #fn
        
        pos = 0
        retstr = ""
        Set Matches = .Execute(str2)   ' 検索を実行します。
        For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
           pos = Match.FirstIndex       '一致する文字列が見つかった位置
           retstr = Match.Value
'           Debug.Print retstr
        Next
        
        If Len(retstr) = 16 Then
            maxRaceNum = CInt(Mid$(retstr, 10, 2))
        Else
            maxRaceNum = CInt(Mid$(retstr, 10, 1))
        End If
        
        For raceNum = 1 To maxRaceNum
            aRaceURL = aResUrl(wA) & Format$(raceNum, "00") & "/"
            ie.Navigate2 aRaceURL
            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop
            
            str = getHTMLString(ie)
            
            str = Replace(str, vbLf, "@")
            
    '        fn = FreeFile
    '        Open src For Output As #fn
    '        Print #fn, str
    '        Close #fn
            
             .Pattern = "複勝\</th\>.+人気.+円.+枠連"
            
            pos = 0
            Set Matches = .Execute(str)   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               resGet = Match.Value
        '       cnt = cnt + 1
        '        ReDim Preserve aResUrl(cnt)
        '        aResUrl(cnt) = "http://keiba.yahoo.co.jp" & Mid$(retstr, 9)
'               Debug.Print raceNum
'               Debug.Print resGet
            Next
            
            resGet = Replace(resGet, ",", "")
            
            '馬番と払い戻しを抽出
            '複勝</th>@<td class="txC resultNo">3</td>@<td>240円</td>@<td class="resultNinki noBdrR"><span>3番人気</span></td>@</tr>@<tr>@<td class="txC resultNo">12</td>@<td>130円</td>@<td class="resultNinki noBdrR"><span>1番人気</span></td>@</tr>@<tr>@<td class="txC resultNo">1</td>@<td>290円</td>@<td class="resultNinki noBdrR"><span>5番人気</span></td>@</tr>@@<tr>@<th class="txC" rowspan="1">枠連
            'resultNo">3</td>@<td>240円
            
             .Pattern = "resultNo""\>[0-9]+\</td\>@\<td\>[0-9]+円"
            
            pos = 0
            cnt2 = -1
            Set Matches = .Execute(resGet)   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                pos = Match.FirstIndex       '一致する文字列が見つかった位置
                retstr = Match.Value
                cnt2 = cnt2 + 1
                ReDim Preserve resMny(cnt2)
                resMny(cnt2) = retstr
            Next
            
            wkPrt = ""
            
            Debug.Print raceNum
            For resLp = 0 To cnt2
                Debug.Print resMny(resLp)
                'resultNo">6</td>@<td>1860円
                pt1 = InStr(resMny(resLp), "</td>")
                pt2 = InStr(resMny(resLp), "<td>")
                wkwk = Mid$(resMny(resLp), 11, pt1 - 11)
                Debug.Print wkwk
                wkwk = Mid$(resMny(resLp), pt2 + 4, Len(resMny(resLp)) - (pt2 + 4))
                Debug.Print wkwk
            Next resLp
            
        Next raceNum
        
    Next wA
    
End With
    

Set objRegExp = Nothing

End Sub

Private Sub Command43_Click()
    Timer1.Enabled = True
End Sub

Private Sub Command44_Click()
    'https://keiba.yahoo.co.jp/odds/tfw/1705040301/?ninki=1
    myURL = "https://keiba.yahoo.co.jp/odds/tfw/1705040301/?ninki=1"
    '起動中のIEを閉じる場合
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
    Set ie = New SHDocVw.InternetExplorer
    '指定のURLを表示
    ie.Navigate2 myURL
    ie.Visible = True    'IE を表示
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Dim str As String
    Dim str2 As String
    Dim aUmaban As String
    Dim aSanrenPuku As String
    Dim aRnk As String
    Dim aTan As String
    Dim aPos As Integer
    
    str = getHTMLString(ie)
    
    aDat = Split(str, vbLf)
    
    For ii = 0 To UBound(aDat)
        If aDat(ii) = "<h3 class=""midashi3rd mgnBS"">枠連</h3>" Then
            Exit For
        End If
        aPos = InStr(aDat(ii), "oddsRank")
        If aPos > 0 Then
            aRnk = Mid$(aDat(ii), aPos + 10)
            aPos = InStr(aRnk, "<")
            aRnk = Left$(aRnk, aPos - 1)
            
            aUmaban = aDat(ii + 1)
            aPos = InStr(aUmaban, "</span></td><td>")
            aUmaban = Mid$(aUmaban, aPos + 16)
            aPos = InStr(aUmaban, "<")
            aUmaban = Left$(aUmaban, aPos - 1)
        End If
    Next ii
    
    '条件に該当していたら、3連複を確認
    'https://keiba.yahoo.co.jp/odds/sf/1705040301/?ninki=1
    myURL = "https://keiba.yahoo.co.jp/odds/sf/1705040301/?ninki=1"
   ie.Navigate2 myURL
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    
    str = getHTMLString(ie)
    
    aDat = Split(str, vbLf)
    
    For ii = 0 To UBound(aDat)
        aPos = InStr(aDat(ii), "oddsRank")
        If aPos > 0 Then
            aPos = InStr(aDat(ii), "class=""txR"">")
            If aPos > 0 Then
                aPos = InStr(aDat(ii), "</td><td>")
                aSanrenPuku = Mid$(aDat(ii), aPos + 9)
                aPos = InStr(aSanrenPuku, "<")
                aSanrenPuku = Left$(aSanrenPuku, aPos - 1)
            End If
        End If
    Next ii
    
End Sub

Private Sub Command45_Click()
    Dim str As String
    'http://www.ipat.jra.go.jp/
    myURL = "https://www.ipat.jra.go.jp/2017/pw_890_i.cgi#!/"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
   ie.Visible = True    'IE を表示
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    str = getHTMLString(ie)
    
    aDat = Split(str, vbLf)
    
End Sub

Private Sub Command46_Click()
   If Not objIE Is Nothing Then
      objIE.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   objIE.Visible = True    'IE を表示
   
    Dim str As String
    Dim aDay As String
    
    aDay = Format$(aStart, "yyyymmdd")
    myURL = "http://localhost/mysql1.php"
    '少佐「イシカワ、HTML取得せよっ！」
    objIE.Navigate2 myURL
    StartTime = GetTickCount
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
        Sleep (100)
        StopTime = GetTickCount
    Loop
    
    Sleep (100)

    str = getHTMLString(objIE)
    
    'イシカワ「電脳に保存しました」
    fn2 = FreeFile
    Open "c:\test\" & aDay & ".dat" For Output As #fn2
    Print #fn2, str
    
    Close #fn2

End Sub

Private Sub Command47_Click()
'http://javpop.com/2012
'http://javpop.com/2014/07/page/2
'http://javpop.com/2014/07/page/45
'Error 404 - Not Found

    Dim str As String
    Dim tmp As Integer
    Dim idxYear As Integer
    Dim idxPage As Integer
    
    If Not ie Is Nothing Then
       ie.Quit
       Set ie = Nothing
    End If
    Set ie = New SHDocVw.InternetExplorer
    ie.Visible = True    'IE を表示
    
    For idxYear = 2017 To 2018
        For idxPage = 1 To 999
             myURL = "http://javpop.com/" & CStr(idxYear) & "/page/" & CStr(idxPage)
            
            ie.Navigate2 myURL
             
             Me.Caption = "Login start"
             Me.Refresh
             
             Do While ie.Busy = True Or ie.ReadyState <> 4
                 DoEvents
             Loop
            
             Me.Caption = "Login comp"
             Me.Refresh
             
            '月のURLを取得する
            str = getHTMLString(ie)
            tmp = InStr(str, "Error 404 - Not Found")
            
            If tmp > 0 Then
                tmp = tmp
                Exit For
            Else
                fnum2 = FreeFile()
                ff = App.Path & "\javpop" & CStr(idxYear) & Format$(idxPage, "000") & ".txt"
                Open ff For Output As #fnum2
                
                Print #fnum2, str
                Close #fnum2
            End If
            
        Next idxPage
    Next idxYear
    
End Sub



Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lstrTmp             As String
    Dim i As Integer
    
On Error GoTo ErrHandler
    
    'ﾄﾞﾛｯﾌﾟされたものが、ﾌｧｲﾙであるか判断
    If Data.GetFormat(vbCFFiles) Then
        For i = 1 To Data.files.Count
            List1.AddItem (Data.files(i))
        Next i
        
    Else
        MsgBox "ドロップされたものがﾌｧｲﾙではありません。"
        Exit Sub
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "error:" & Err.Description
    Exit Sub

End Sub


' IEを閉じた時のQuitイベントをフックする
Private Sub objIE_OnQuit()
''    ' Excelの画面上にメッセージを表示する
''    MsgBox "IEを閉じました"
''    Set objIE = Nothing
End Sub

Private Sub objIE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim str As String
    
    ' Excelの画面上にメッセージを表示する
    document_completed_flag = True
    Debug.Print URL
    mURL = URL
    
    If InStr(URL, "http://keiba.rakuten.co.jp/odds") > 0 Then
        str = getHTMLString(pDisp)
        Dim fn2  As Long
        fn2 = FreeFile
        Open "c:\0210.html" For Output As #fn2
        Print #fn2, str
        
        Close #fn2
    End If
    
End Sub

Private Sub Command2_Click()
    myURL = "https://id.nikkansports.com/u/member/login/?guid=on&cid=53&premium=true&backurl=http://p.nikkansports.com/premium%2fj_spring_nikkan_security_check%3Fcurl%3dhttp%253A%252F%252Fp%2enikkansports%2ecom%252Fgoku-uma%252Fmember%252Findex%2ezpl&level=1"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
    If chkD.Value = 1 And chkDL.Value = 0 Then
        ie.Visible = True    'IE を表示
    End If
    
    Me.Caption = "Login start"
    Me.Refresh
    
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Me.Caption = "Login comp"
    Me.Refresh

End Sub

Private Sub Command3_Click()
    Call IE_DocumentComplete(ie, myURL)
End Sub

Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)

   With pDisp
      .Document.Forms(0).elements("mailAddress").Value = "jun@buhi-buhi.com"
      .Document.Forms(0).elements("password").Value = "two784"
      
'Sleep (10)
      
'      .Document.getElementsByTagName("INPUT")(2).submit
'      .Document.Forms(0).elements(2).Click
      .Document.getElementsByTagName("INPUT")(9).Click
   End With
    
'Dim objForm As HTMLFormElement    'page_member_login_MemberLoginForm
'    Set objForm = pDisp.Document.Forms("page_member_login_MemberLoginForm")
'    objForm.submit
    
'    pDisp.Navigate2 "view-source:https://id.nikkansports.com/u/member/login/MemberLogin.do"
    
'    'キャッシュは消すこと
'    'http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2016&mode=kako
'
    Me.Caption = "Login input start"
    Me.Refresh
    
    Do While pDisp.Busy = True Or pDisp.ReadyState <> 4
        DoEvents
    Loop

    Me.Caption = "Login input comp"
    Me.Refresh
''
'''    pDisp.Navigate2 "http://p.nikkansports.com/goku-uma/member/compi/compi_list.zpl?year=2016&mode=kako"
'''''''''    pDisp.Navigate2 "http://p.nikkansports.com/goku-uma/member/compi/compi.zpl?course_id=006&date=20160105"
'
'    Do While pDisp.Busy = True Or pDisp.ReadyState <> 4
'        DoEvents
'    Loop
'
''    Call SendKeys("%fa")
    
'    Call Sleep(1)
'
'    Do While pDisp.Busy = True Or pDisp.ReadyState <> 4
'        DoEvents
'    Loop
'
'
'
'    Call SendKeys("%f")
'    Call SendKeys("a")
End Sub

Private Function getHTMLString(ie As Object) As String
On Error GoTo err_handler
    Dim htdoc As HTMLDocument
    Set htdoc = ie.Document
    
    Dim ret As String
    ret = htdoc.getElementsByTagName("HTML")(0).outerHTML & vbCrLf
    getHTMLString = ret
    
    Exit Function
err_handler:
    
    Debug.Print Err.Description
    
    getHTMLString = ""
    
    Exit Function
End Function

Private Sub Command4_Click()
    Dim str As String
    
    str = getHTMLString(ie)

Debug.Print str

End Sub

Private Sub Command5_Click()
    'http://keiba.yahoo.co.jp/schedule/list/2007/?month=10
    
    
    myURL = "http://keiba.yahoo.co.jp/schedule/list/" & txtY.Text & "/?month=1"
   '起動中のIEを閉じる場合
   If Not ie Is Nothing Then
      ie.Quit
      Set ie = Nothing
   End If
   Set ie = New SHDocVw.InternetExplorer
   '指定のURLを表示
   ie.Navigate2 myURL
'   ie.Visible = True    'IE を表示
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop

    Dim str As String
    Dim str2 As String
    
    str = getHTMLString(ie)

'    Debug.Print str
    
    
    Dim fn As Long
    Dim src As String
    
    src = "c:\temp\ya.txt"
'    fn = FreeFile
'    Open src For Output As #fn
'    Print #fn, str
'    Close #fn
    
Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
Dim strResult As String '置換後の文字列
Dim Matches
Dim Match
    Dim aResUrl() As String
    Dim cnt As Long
    Dim cnt2 As Long
    Dim wA As Integer
    
    Dim raceNum As Integer
    Dim maxRaceNum As Integer
    Dim aRaceURL As String
    Dim lstUrl As String
    Dim resGet As String
    Dim resLp As Integer
    Dim resMny() As String
    Dim wkPrt As String
    Dim pt1 As Integer
    Dim pt2 As Integer
    Dim wkwk As String
    
    cnt = -1
    
'<<データ解析>>
'正規表現オブジェクトの宣言
Set objRegExp = New RegExp

With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True
    
    'year loop
    
    'month loop
    
    '各年毎の各月のデータ取得
     .Pattern = "a href=""" & "/" & "race" & "/" & "result" & "/........../"
    
    pos = 0
    Set Matches = .Execute(str)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
        ReDim Preserve aResUrl(cnt)
        aResUrl(cnt) = "http://keiba.yahoo.co.jp" & Mid$(retstr, 9)
        aResUrl(cnt) = Left$(aResUrl(cnt), Len(aResUrl(cnt)) - 3)
'       Debug.Print retstr
    Next
    
    raceNum = 1
    
    For wA = 0 To cnt
        '任意の日、開催場所のHTMLから、各レース毎のURLを取得
        
        'レースMax.を取得
        lstUrl = Left$(aResUrl(wA), 30) & "list/" & Right$(aResUrl(wA), 9)
        ie.Navigate2 lstUrl
'        Do While ie.Busy = True Or ie.ReadyState <> 4
'            Call Sleep(1)
'            DoEvents
'        Loop
        Do While ie.Busy = True Or ie.ReadyState <> 4
'            Call Sleep(1)
'            If ie.Busy = False Then
'                Exit Do
'            End If
'            If ie.ReadyState = READYSTATE_COMPLETE Then
'                Exit Do
'            End If
            
            DoEvents
        Loop
        
        'scheRNo">.+R</p>
         .Pattern = "scheRNo""\>.+R\</p\>"
        str2 = ""
        str2 = getHTMLString(ie)
'    fn = FreeFile
'    Open src For Output As #fn
'    Print #fn, str2
'    Close #fn
        
        pos = 0
        retstr = ""
        Set Matches = .Execute(str2)   ' 検索を実行します。
        For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
           pos = Match.FirstIndex       '一致する文字列が見つかった位置
           retstr = Match.Value
'           Debug.Print retstr
        Next
        
        If Len(retstr) = 16 Then
            maxRaceNum = CInt(Mid$(retstr, 10, 2))
        Else
            maxRaceNum = CInt(Mid$(retstr, 10, 1))
        End If
        
        For raceNum = 1 To maxRaceNum
            aRaceURL = aResUrl(wA) & Format$(raceNum, "00") & "/"
            ie.Navigate2 aRaceURL
            Do While ie.Busy = True Or ie.ReadyState <> 4
                DoEvents
            Loop
            
            str = getHTMLString(ie)
            
            str = Replace(str, vbLf, "@")
            
    '        fn = FreeFile
    '        Open src For Output As #fn
    '        Print #fn, str
    '        Close #fn
            
             .Pattern = "複勝\</th\>.+人気.+円.+枠連"
            
            pos = 0
            Set Matches = .Execute(str)   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
               pos = Match.FirstIndex       '一致する文字列が見つかった位置
               resGet = Match.Value
        '       cnt = cnt + 1
        '        ReDim Preserve aResUrl(cnt)
        '        aResUrl(cnt) = "http://keiba.yahoo.co.jp" & Mid$(retstr, 9)
'               Debug.Print raceNum
'               Debug.Print resGet
            Next
            
            resGet = Replace(resGet, ",", "")
            
            '馬番と払い戻しを抽出
            '複勝</th>@<td class="txC resultNo">3</td>@<td>240円</td>@<td class="resultNinki noBdrR"><span>3番人気</span></td>@</tr>@<tr>@<td class="txC resultNo">12</td>@<td>130円</td>@<td class="resultNinki noBdrR"><span>1番人気</span></td>@</tr>@<tr>@<td class="txC resultNo">1</td>@<td>290円</td>@<td class="resultNinki noBdrR"><span>5番人気</span></td>@</tr>@@<tr>@<th class="txC" rowspan="1">枠連
            'resultNo">3</td>@<td>240円
            
             .Pattern = "resultNo""\>[0-9]+\</td\>@\<td\>[0-9]+円"
            
            pos = 0
            cnt2 = -1
            Set Matches = .Execute(resGet)   ' 検索を実行します。
            For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
                pos = Match.FirstIndex       '一致する文字列が見つかった位置
                retstr = Match.Value
                cnt2 = cnt2 + 1
                ReDim Preserve resMny(cnt2)
                resMny(cnt2) = retstr
            Next
            
            wkPrt = ""
            
            Debug.Print raceNum
            For resLp = 0 To cnt2
                Debug.Print resMny(resLp)
                'resultNo">6</td>@<td>1860円
                pt1 = InStr(resMny(resLp), "</td>")
                pt2 = InStr(resMny(resLp), "<td>")
                wkwk = Mid$(resMny(resLp), 11, pt1 - 11)
                Debug.Print wkwk
                wkwk = Mid$(resMny(resLp), pt2 + 4, Len(resMny(resLp)) - (pt2 + 4))
                Debug.Print wkwk
            Next resLp
            
        Next raceNum
        
    Next wA
    
End With
    

Set objRegExp = Nothing
    
    
    
End Sub

Private Sub Command6_Click()
   '起動中のIEを閉じる場合
   If Not objIE Is Nothing Then
      ie.Quit
      Set objIE = Nothing
   End If
'   Set objIE = New SHDocVw.InternetExplorer
    Set objIE = CreateObject("InternetExplorer.Application")
   
   
    Dim str As String
    
   
   '指定のURLを表示
    myURL = "https://bet.keiba.rakuten.co.jp/bank/deposit/"
   objIE.Navigate2 myURL
   objIE.Visible = True    'IE を表示
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop

    Sleep (100)
    
    str = getHTMLString(objIE)
   
   If InStr(str, "楽天会員ログイン") > 0 Then
        'login form
        
        str = str
        objIE.Document.All("u").Value = "granbri@gmail.com"
        objIE.Document.All("p").Value = "two784jun"
        objIE.Document.All("submit").Click
        
        Do While objIE.Busy = True Or objIE.ReadyState <> 4
            DoEvents
        Loop
        
   End If
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\ya.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn
    
    objIE.Document.All("price").Value = "100"

'    objIE.Document.getElementsByTagName("A")(8).Click
    objIE.Document.All("depositingInputButton").Click
'    objIE.Document.getElementsByTagName("depositingInputButton")(0).Click

    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop

    Sleep (100)
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\" & Format$(Now, "hhmmss") & "betTest.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn
    
'    objIE.Document.getElementsByTagName("A")(9).Click
    
    Do
        DoEvents
        If mURL = "https://bet.keiba.rakuten.co.jp/bank/deposit/confirm" Then
            Exit Do
        End If
    Loop
    
    
    If InStr(str, "<div class=""errorMsg"" id=""depositingConfirmValidMessage"" style=""display:none;"">暗証番号を入力してください。</div>") = 0 Then
        objIE.Document.All("pin").Value = "0358"
    End If
'    objIE.Document.All("pin").Value = "0358"
    objIE.Document.All("depositingConfirmButton").Click



    Sleep (1000)






    myURL = "https://bet.keiba.rakuten.co.jp/bet_lite/"
   objIE.Navigate2 myURL
   objIE.Visible = True    'objIE を表示
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop

    Sleep (100)
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\ya.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn

    '<option value="20" >大井</option>          '他場所のデータを確認すること
    '<option value="24" >名古屋</option>
    '<option value="27" >園田</option>
    '<option value="31" >高知</option>
    '浦和   18
    '笠松   32
    '
    
    objIE.Document.All("keibajouCode").Value = 31
    objIE.Document.All("raceNumber").Value = txtRace.Text
    objIE.Document.All("shikibetsu").Value = 2
    objIE.Document.All("houshiki").Value = 16
    objIE.Document.getElementsByTagName("INPUT")(6).Click

    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop

    Sleep (100)
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\ya.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn
    
    'objIE.Document.getElementsByTagName("radio")(5).Click
'    objIE.Document.getElementsByName("me1[]")(3).Click     '4番
    objIE.Document.getElementsByName("me1[]")(CInt(txtUma.Text) - 1).Click   '(n) n+1が、馬番
    objIE.Document.All("buyUnitCount").Value = 1
'    objIE.Document.getElementsByTagName("INPUT")(17).Click     '馬番が６までなら、１７。つまり、最大馬番－１１
    objIE.Document.All("confirm").Click
    
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop

    Sleep (100)
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\ya.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn

    Sleep (100)
    ' 投票Lite:    投票内容確認
    If InStr(str, "--><p class=""codeArea""><input name=""passcode") > 0 Then
        objIE.Document.All("passcode").Value = "0358"
    End If
    objIE.Document.All("cashConfirm").Value = 100
    
'    objIE.Document.getElementsByTagName("INPUT")(14).Click '9は、買い目を追加
'    objIE.Document.getElementsByName("vote")(0).Click
'    objIE.Document.getElementsByName("add")(0).Click
'    objIE.Document.getElementsByTagName("INPUT")(9).Click '9は、買い目を追加
   objIE.Navigate2 "JavaScript: document.frmVote.submit()"
    
    Do While objIE.Busy = True Or objIE.ReadyState <> 4
        DoEvents
    Loop
    
    Sleep (100)
    
    str = getHTMLString(objIE)
    
    src = "c:\temp\ya.txt"
    fn = FreeFile
    Open src For Output As #fn
    Print #fn, str
    Close #fn
    
    '引き続き投票する
    objIE.Document.getElementsByName("top")(0).Click
    
End Sub

Private Sub Command7_Click()
    Call getNankanCmpiList
    
'    Dim str As String
'    Dim aYear As Integer
'    Dim aDay As Integer
'    Dim aRace As Integer
'    Dim aUma As Integer
'    Dim aGatu As String
'    Dim aNiti As String
'    Dim aYmd As String
'    Dim prt As String
'    Dim dbg As String
'    Dim timenow As String
'
'    Dim fnum As Integer
'    fnum = FreeFile()
'
'    timenow = Format$(Now, "hh:mm:ss")
'
'    Open "c:\temp\daily\nankan-" & areaY.Text & areaMD(0).Text & areaMD(1).Text & "-" & Format$(Now, "yyyymmddhhmmss") & ".txt" For Output As #fnum
'
'    Me.Caption = "start"
'    Me.Refresh
'    '指定のURLを表示
'    myURL = "http://p.nikkansports.com/goku-uma/member/races/past_list_nankan.zpl"
'
'    ie.Navigate2 myURL
'
'    Do While ie.Busy = True Or ie.ReadyState <> 4
'        DoEvents
'    Loop
'    str = getHTMLString(ie)
'
'    Me.Caption = "comp"
'    Me.Refresh
'
'    '年のURLを取得する
'    If str = "" Then
'        GoTo exit_here
'    End If
'    gStr = str
'    Call getYear(1)
'
'    '年ループ   gYear gUrlYear
'    For aYear = 0 To UBound(gYear)
'        If gYear(aYear) = areaY.Text Or areaY.Text = "" Then
'            '指定の年サイトに移動
'            myURL = gUrlYear(aYear)
'
'            Me.Caption = "start"
'            Me.Refresh
'            ie.Navigate2 myURL
'
'            Do While ie.Busy = True Or ie.ReadyState <> 4
'                DoEvents
'            Loop
'            str = getHTMLString(ie)
'
'            Me.Caption = "comp"
'            Me.Refresh
'
'            'すべての日付のURLを取得する
'            If str = "" Then
'                GoTo exit_here
'            End If
'            gStr = str
'            Call getDay(1)
'
'            '日付ループ gDay gPosDay
'            For aDay = 0 To UBound(gDay)
'                If (gDayFmt(aDay) >= areaMD(0).Text And gDayFmt(aDay) <= areaMD(1).Text) Or ("" = areaMD(0).Text And "" = areaMD(1).Text) Then
'                    aGatu = Mid$(gDay(aDay), 1, InStr(gDay(aDay), "月") - 1)
'                    aNiti = Mid$(gDay(aDay), InStr(gDay(aDay), "月") + 1)
'                    aNiti = Left$(aNiti, Len(aNiti) - 1)
'                    aYmd = gYear(aYear) & Format$(aGatu, "00") & Format$(aNiti, "00")
'                    'If Format$(Now, "yyyymmdd") > aYmd Then
'                        '任意の日付
'
'                        Me.Caption = "start"
'                        Me.Refresh
'
'                        myURL = gUrlDay(aDay)
'                        ie.Navigate2 myURL
''                        ie.Visible = True    'IE を表示
'                        Do While ie.Busy = True Or ie.ReadyState <> 4
'                            DoEvents
'                        Loop
'
'                        Me.Caption = "comp"
'                        Me.Refresh
'
'                        '全レースのURLを取得する
'                        str = getHTMLString(ie)
'                        If str = "" Then
'                            GoTo exit_here
'                        End If
'                        gStr = str
'                        Call getRaces
'
'                        'コンピ指数
'                        Me.Caption = "start"
'                        Me.Refresh
'
'                        myURL = gCmpDay(0)
'                        ie.Navigate2 myURL
''                        ie.Visible = True    'IE を表示
'                        Do While ie.Busy = True Or ie.ReadyState <> 4
'                            DoEvents
'                        Loop
'
'                        Me.Caption = "comp"
'                        Me.Refresh
'
'                        'コンピ指数 ファイル保存 gCmpDay
'                        str = getHTMLString(ie)
'                        If str = "" Then
'                            GoTo exit_here
'                        End If
'                        gStr = str
'
'                        'レースループ
'                        For aRace = 0 To UBound(gDenmaRace)
'        '                    myURL = gUrlDay(aDay)
'        '                    ie.Navigate2 myURL
'        '                    ie.Visible = True    'IE を表示
'        '                    Do While ie.Busy = True Or ie.ReadyState <> 4
'        '                        DoEvents
'        '                    Loop
'
'                            '出走表
'                            myURL = gDenmaRace(aRace)
'
'                            Me.Caption = "start"
'                            Me.Refresh
'
'                            If chkDL.Value = 0 Then
'                                ie.Navigate2 myURL
'                            Else
'                                ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                                DoEvents
'                            End If
'
'                            If chkDL.Value = 0 Then
'                                Do While ie.Busy = True Or ie.ReadyState <> 4
'                                    DoEvents
'                                Loop
'                                str = getHTMLString(ie)
'                            Else
'                                Call TextCodeChg(SaveFileName)
'                                str = getHtmlFile
'                            End If
'
'                            Me.Caption = "comp"
'                            Me.Refresh
'
'                            '出走表 すべての馬の馬番と馬名を取得する gBamei gUmaban
'                            If str = "" Then
'                                GoTo exit_here
'                            End If
'                            gStr = str
'                            Call getRunTable
'
'        '                    myURL = gUrlDay(aDay)
'        '                    ie.Navigate2 myURL
'        '                    ie.Visible = True    'IE を表示
'        '                    Do While ie.Busy = True Or ie.ReadyState <> 4
'        '                        DoEvents
'        '                    Loop
'
'
'                            '年月日(gYear(aYear) & gDay(aDay))、開催場所(gPosDay(aDay))、レース番号(gRace(aRace))、馬番、馬名(gBamei gUmaban)をファイルに出力する
'                            Debug.Print gYear(aYear) & "," & gDay(aDay) & "," & gRace(aRace)
'                            For aUma = 0 To UBound(gUmaban)
'                                'Debug.Print gUmaban(aUma) & "," & gBamei(aUma)
'
'                                prt = "1," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gUmaban(aUma) & "," & gBamei(aUma) & "," & gCmp(aUma)
'                                Debug.Print prt
'                                Print #fnum, prt
'                            Next aUma
'
'                            '結果
'                            If Format$(Now, "yyyymmdd") > aYmd Then
'                                If UBound(gResRace) >= aRace Then
'                                    myURL = gResRace(aRace)
'
'                                    Me.Caption = "start"
'                                    Me.Refresh
'
'                                    If chkDL.Value = 0 Then
'                                        ie.Navigate2 myURL
'                                    Else
'                                        ret = URLDownloadToFile(0, myURL, SaveFileName, 0, 0)
'                                        DoEvents
'                                    End If
'
'                                    If chkDL.Value = 0 Then
'                                        Do While ie.Busy = True Or ie.ReadyState <> 4
'                                            DoEvents
'                                        Loop
'                                        str = getHTMLString(ie)
'                                    Else
'                                        Call TextCodeChg(SaveFileName)
'                                        str = getHtmlFile
'                                    End If
'
'                                    Me.Caption = "comp"
'                                    Me.Refresh
'
'                                    '結果 とりあえず、複勝のみ gFukuMny gFukuNum
'                                    If str = "" Then
'                                        GoTo exit_here
'                                    End If
'                                    gStr = str
'                                    Call getRes
'
'                                    '結果(gFukuMny gFukuNum)をファイルに出力する
'                                    For aUma = 0 To UBound(gFukuNum)
'                                        Debug.Print gFukuNum(aUma) & "," & gFukuMny(aUma)
'                                        prt = "2," & gYear(aYear) & "," & gDay(aDay) & "," & gDayFmt(aDay) & "," & gPosDayCd(aDay) & "," & gPosDayDbCd(aDay) & "," & gRace(aRace) & "," & gFukuNum(aUma) & "," & gFukuMny(aUma)
'                                        Debug.Print prt
'                                        Print #fnum, prt
'                                    Next aUma
'                                End If
'                            End If
'
'                            prt = prt
'                        Next aRace
'                    'End If
'                End If
'            Next aDay
'        End If
'    Next aYear
'
'exit_here:
'    Close #fnum
'
'    Debug.Print "start:" & timenow
'    Debug.Print "end  :" & Format$(Now, "hh:mm:ss")
    
End Sub

Private Sub Command8_Click()
    src = "c:\temp\denma.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

'     .Pattern = "馬番 -->.+/span"
     .Pattern = "馬番 -->*" & kaigyo & ".+td>" & kaigyo & ".+性齢取得 -->" & kaigyo & ".+コンピ -->" & kaigyo & ".+/span"
'     .Pattern = "馬番 -->*" & kaigyo & ".+td>"
    
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
       Debug.Print retstr
    Next
    
End With
    
End Sub

'結果get
Private Sub getRes()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
'    kaigyo = vbCr & "$" & vbLf & "^"
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    .Pattern = "複勝.+円"
    cnt = -1
    
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
       
'       Debug.Print retstr
    
        ReDim Preserve gFukuMny(cnt)
        ReDim Preserve gFukuNum(cnt)
        
        If chkDL.Value = 0 Then
            gWk = Mid$(retstr, 12)
            gWk = Left$(gWk, InStr(gWk, "</td") - 1)
            gFukuNum(cnt) = gWk
            gWk = Mid$(retstr, 12 + Len(gFukuNum(cnt)) + 9)
            gWk = Left$(gWk, Len(gWk) - 1)
            gFukuMny(cnt) = Replace(gWk, ",", "")
        Else
            gWk = Mid$(retstr, 21)
            gWk = Left$(gWk, InStr(gWk, "</td") - 1)
            gFukuNum(cnt) = gWk
            gWk = Mid$(retstr, 12 + Len(gFukuNum(cnt)) + 27)
            gWk = Left$(gWk, Len(gWk) - 1)
            gFukuMny(cnt) = Replace(gWk, ",", "")
            gFukuNum(cnt) = Format$(gFukuNum(cnt), "00")
        End If
        
        
'        Debug.Print gFukuMny(cnt)
    Next
    
End With

End Sub

'中央競馬の結果
Private Sub getChuuouRes()
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
'    kaigyo = vbCr & "$" & vbLf & "^"
    kaigyo = vbLf
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    .Pattern = "複勝.+円"
    cnt = -1
    
    pos = 0
    Set Matches = .Execute(gStr)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
       
'       Debug.Print retstr
    
        ReDim Preserve gFukuMny(cnt)
        ReDim Preserve gFukuNum(cnt)
        
        gWk = Mid$(retstr, 12)
        gWk = Left$(gWk, InStr(gWk, "</td") - 1)
        gFukuNum(cnt) = gWk
'        Debug.Print gWk
        gWk = Mid$(retstr, 12 + Len(gFukuNum(cnt)) + 9)
        gWk = Left$(gWk, Len(gWk) - 1)
        gFukuMny(cnt) = Replace(gWk, ",", "")
        
'        Debug.Print gFukuMny(cnt)
    Next
    
End With

End Sub

'結果get
Private Sub Command9_Click()
    src = "c:\temp\result.txt"
    
    fn = FreeFile
    Open src For Input As #fn
    
    '<<ファイル 読>>
    lCnt = 0
    Line Input #fn, wk
    wkall = wk
    
    Do Until EOF(fn)
        Line Input #fn, wk
        wkall = wkall & vbCr & vbLf & wk
    Loop
    
    '<<ファイル 閉>>
    Close #fn
    
    Dim objRegExp As RegExp 'RegExp：[参照設定]で Microsoft VBScript Regular Expressions 5.5 にチェックを付ける
    Dim strResult As String '置換後の文字列
    Dim Matches
    Dim Match
    '正規表現オブジェクトの宣言
    Set objRegExp = New RegExp
    
    Dim kaigyo As String
    
    kaigyo = vbCr & "$" & vbLf & "^"
    
With objRegExp
    .Global = True '複数マッチ可
    .IgnoreCase = True
    .Global = True
    .MultiLine = True

    .Pattern = "複勝.+円"
    cnt = -1
    
    pos = 0
    Set Matches = .Execute(wkall)   ' 検索を実行します。
    For Each Match In Matches   ' Matches コレクションに対して繰り返し処理を行います。
       pos = Match.FirstIndex       '一致する文字列が見つかった位置
       retstr = Match.Value
       cnt = cnt + 1
       
'       Debug.Print retstr
    
        ReDim Preserve gFukuMny(cnt)
        ReDim Preserve gFukuNum(cnt)
        
        gWk = Mid$(retstr, 21)
        gWk = Left$(gWk, InStr(gWk, "/td") - 2)
        gFukuNum(cnt) = gWk
'        Debug.Print gWk
        gWk = Mid$(retstr, 21)
        gWk = Mid$(gWk, InStr(retstr, "class=") + 8)
        gWk = Left$(gWk, Len(gWk) - 1)
        gWk = Replace(gWk, ",", "")
        gFukuMny(cnt) = gWk
        
        Debug.Print gFukuMny(cnt)
    Next
    
End With

End Sub

Private Sub Form_Load()
'''    Me.Left = 0
'''    Me.Top = 0
'''
'''    'MACアドレス取得
'''    Dim aMac As String
'''    aMac = getMacAddress
'''
'''    'Web情報と照合
'''    Dim SaveFileName As String
'''    Dim DownloadFile As String
'''    Dim ret          As Long
'''    Dim str As String
'''
'''    SaveFileName = App.Path & "\0519.txt"
'''    DownloadFile = "http://buhi-buhi.com/apl/mac.dat"
'''
'''    ret = URLDownloadToFile(0, DownloadFile, SaveFileName, 0, 0)
'''    DoEvents
'''
'''    fnum = FreeFile()
'''
'''    Open SaveFileName For Input As #fnum
'''
'''    Do Until EOF(fnum)
'''        Line Input #fnum, wk
'''        str = str & vbLf & wk
'''    Loop
'''
'''    Close #fnum
'''
'''    Kill (SaveFileName)
'''
'''    Dim arr() As String
'''    Dim aDat() As String
'''    Dim aFlg As Boolean
'''    Dim aID As String
'''
'''    aFlg = False
'''    arr = Split(str, vbLf)
'''
'''    '登録存在チェック
'''    For ii = 1 To UBound(arr)
'''        aDat = Split(arr(ii), ",")
'''        If aDat(1) = aMac Then
'''            aID = aDat(0)
'''            aFlg = True
'''            Exit For
'''        End If
'''    Next ii
'''
'''    If aMac <> ",5C:51:4F:8F:66:BD,CC:7E:E7:5F:AD:8B,5C:51:4F:8F:66:B9,5E:51:4F:8F:66:BA,5E:51:4F:8F:66:B9,78:61:7C:C1:54:13" Then
'''
'''
'''        Dim mail As String
'''        Dim aTitle As String
'''        Dim aBody As String
'''
'''        If aFlg = True Then
'''rrn_here:
'''            'ID入力
'''            rtn = InputBox("IDを入力して下さい", "OKVA")
'''            If StrPtr(rtn) = 0 Then
'''                MsgBox "キャンセルが選択されました"
'''                End
'''            Else
'''                If rtn = "" Then
'''                    MsgBox "文字が入力されていません"
'''                    End
'''                Else
'''                    If rtn = aID Then
'''                        MsgBox "ID確認できました。"
'''                    Else
'''                        MsgBox "ID確認できませんでした。"
'''                        GoTo rrn_here
'''                    End If
'''
'''                End If
'''            End If
'''        Else
'''            'メールアドレス入力
'''            rtn = InputBox("メールアドレスを入力して下さい", "OKVA")
'''            If StrPtr(rtn) = 0 Then
'''                MsgBox "キャンセルが選択されました"
'''                End
'''            Else
'''                If rtn = "" Then
'''                    MsgBox "文字が入力されていません"
'''                    End
'''                Else
'''                    MsgBox "登録処理をいたします。でき次第、メールにて連絡いたします。"
'''
'''                    aTitle = "○ソフトウェア登録処理希望"
'''                    aBody = rtn & "," & aMac
'''
'''                    mail = sendMail(aTitle, aBody, "regist@buhi-buhi.com")
'''                    MsgBox "ご連絡をお待ちくださいませ！"
'''
'''                    End
'''                End If
'''            End If
'''        End If
'''
'''        aTitle = "起動：オーケー馬"
'''        aBody = rtn & "," & aMac
'''
'''        mail = sendMail(aTitle, aBody, "racesoft@buhi-buhi.com")
'''    End If
'''
'''    SaveFileName = "C:\temp\xxx.htm"
'''
'''    If CreateObject("NonCodeVb6.NonCodeClass") Is Nothing Then
'''        If Len(Dir("NonCodeVb6.dll")) <> 0 Then
'''            ' NonCodeVb6.dllのレジストリ登録
'''            Shell "regsvr32 /s NonCodeVb6.dll", vbHide
'''        Else
'''            ' NonCodeVb6.dllをCode2Code.exeと同じフォルダに置いてください。
'''            MsgBox _
'''            "NonCodeVb6.dllが見つかりませんでした。" & vbCrLf & vbCrLf & _
'''            "NonCodeVb6.dllを" & vbCrLf & "[" & App.Path & "]" & vbCrLf & _
'''            "に置いてください。"
'''            End
'''        End If
'''    End If
'''
'''    Set objNonCode = CreateObject("NonCodeVb6.NonCodeClass")
'''
'''    areaMD(0).Text = Format$(Now, "mmdd")
'''    areaMD(1).Text = Format$(Now, "mmdd")
    
    Set objNonCode = CreateObject("NonCodeVb6.NonCodeClass")
    
    'DB接続
'    gRet = cnctDB
    
End Sub

Private Function sendMail(msg_subject As String, msg_body As String, aite As String) As String
On Error GoTo err_handler

Set objMail = CreateObject("CDO.Message")

objMail.From = "o.k.keiba@gmail.com"
objMail.To = aite
objMail.Subject = msg_subject
objMail.HTMLBody = msg_body

Dim strConfigurationField  As String

strConfigurationField = "http://schemas.microsoft.com/cdo/configuration/"
With objMail.Configuration.Fields
   .Item(strConfigurationField & "sendusing") = 2
   .Item(strConfigurationField & "smtpserver") = "smtp.googlemail.com"
   .Item(strConfigurationField & "smtpserverport") = 465
   .Item(strConfigurationField & "smtpusessl") = True
   .Item(strConfigurationField & "smtpauthenticate") = 1
   .Item(strConfigurationField & "sendusername") = "o.k.keiba@gmail.com"
   .Item(strConfigurationField & "sendpassword") = "lets.keiba7"
   .Item(strConfigurationField & "smtpconnectiontimeout") = 60
   .Update
End With

objMail.send

Set objMail = Nothing

sendMail = ""

Exit Function

err_handler:
    
sendMail = Err.Description

End Function

Public Function getMacAddress() As String

    Dim objNetwork As Object 'Windowsの情報
    Dim strNetworkSql As String 'Windowsの情報取得文 保存変数
    Dim strMacAdr As String '取得したMACアドレス文字
    
    'Windowsの情報取得文 組み立て
    strNetworkSql = "SELECT * FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL"
    
    'Windowsの情報取得文を使い情報を取得(1個目のみ)
    For Each objNetwork In GetObject("winmgmts:").ExecQuery(strNetworkSql)
        strMacAdr = strMacAdr & "," & objNetwork.MACAddress
'        Exit For
    Next
    
    'メッセージボックスへMACアドレスを表示
    getMacAddress = strMacAdr

End Function

Private Sub Command26_Click()
    Dim aaa As String
    Dim cnt As Long
    Dim ii As Long
    Dim dd() As String
    
    aaa = Text1.Text
    
    Call TextCodeChg(aaa)
    
    fn = FreeFile
    Open aaa For Input As #fn
    
'    Do Until EOF(fn)
        Line Input #fn, wk
'        cnt = cnt + 1
'    Loop
    
    Close #fn
  
  Dim Stream As Object
  
  ' VB標準のADODB.Streamオブジェクトを作成する
  Set Stream = CreateObject("ADODB.Stream")
  
  ' ストリームの文字コードをUTF8に設定する
  Stream.Charset = "UTF-8"
  ' ファイルのタイプ(1:バイナリ 2:テキスト)
  Stream.Type = 2
  ' ストリームを開く
  Stream.Open
  ' ストリームの保存形式をテキスト形式にする
  Stream.WriteText wk
  ' ストリームに名前を付けて保存する(1は新規作成 2は上書き保存)
  Stream.SaveToFile (aaa & "_ex.txt"), 2
  ' ストリームを閉じる
  Stream.Close
  
  
'Dim buffer As String
'
'' ストリームを開く
'  Stream.Open
'  ' ストリームにファイルを読み込む
'  Stream.LoadFromFile (aaa & "_ex.txt")
'  ' ファイルの中身をbufferへ代入
'  buffer = Stream.ReadText
'  ' ストリームを閉じる
'  Stream.Close
'
'  ' イミディエイトへ出力
'  Debug.Print buffer
'
'  Set Stream = Nothing
'
'    dd = Split(buffer, vbLf)
'
'    fn = FreeFile
'    Open aaa & "_exz.txt" For Output As #fn
'
'    For ii = 0 To UBound(dd)
'        Print #fn, dd(ii)
'    Next ii
'
'    Close #fn
End Sub

Private Sub Timer1_Timer()
    Dim aMin As Integer
    Dim ii As Integer
    Dim jj As Integer
    Dim aDat() As String
    Dim aHassoTime As String
    Dim aBasicParam() As String
    Dim aa As Variant
    Dim aChkJyo() As String
    Dim aChkRace() As String
    Dim aChkUmaban() As String
    Dim aChkHimo() As String
    Dim aChkFlg As Boolean
    Dim aRet As Integer
    Dim aStr As String
    Dim aChkTarget As Integer
    
    aChkJyo = Split("05,08", ",")
    aChkRace = Split("01,07", ",")
    aChkUmaban = Split("07,09", ",")
    aChkHimo = Split("02-04-05-06-10-12,05-06-07", ",")
    
    'aBasicDat(aBasicCnt) = aJyoCode & "," & aRaceNum & "," & aHassoTime & "," & aJyoLink
    
    aMin = -1 * CInt(Text3.Text)
    
    For ii = 0 To UBound(aBasicDat)
        aBasicParam = Split(aBasicDat(ii), ",")
        'aJyoCode & "," & aRaceNum & "," & aHassoTime & "," & aJyoLink
        aDat = Split(aBasicDat(ii), ",")
        
        'チェックするレースなら
        aChkFlg = False
        For jj = 0 To UBound(aChkJyo)
            If aDat(0) = aChkJyo(jj) And Format$(aDat(1), "00") = aChkRace(jj) Then
                aChkTarget = jj
                aChkFlg = True
                Exit For
            End If
        Next jj
        If aChkFlg = True Then
            '発走時刻n分前？
            aa = aCheckTime(aBasicParam(2), aMin)
            
            If aa < 0 Then      'aa > 0
                'https://keiba.yahoo.co.jp/odds/tfw/1708040201/
                myURL = "https://keiba.yahoo.co.jp/odds/tfw/" & aBasicParam(3) & Format$(aBasicParam(1), "00") & "/"
                ie.Navigate2 myURL
                Do While ie.Busy = True Or ie.ReadyState <> 4
                    DoEvents
                Loop
                
                aStr = getHTMLString(ie)
                
                aDat = Split(aStr, vbLf)
                
'                '発走時刻の確認
'                For jj = 0 To UBound(aDat)
'                    If jj = 174 Then
'                        jj = jj
'                    End If
'                    aWkStr = "発走"
'                    aWkPos = InStr(aDat(jj), aWkStr)
'                    If aWkPos > 0 Then
'                        aHassoTime = Mid$(aDat(jj), aWkPos - 5, 5) 'time
'                        If aBasicParam(2) <> aHassoTime Then
'                            aBasicDat(ii) = aBasicParam(0) & "," & aBasicParam(1) & "," & aHassoTime & "," & aBasicParam(3)
'                        End If
'                        Exit For
'                    End If
'                Next jj
                
                '単勝人気の確認
                aRet = ChkTanSan(aBasicParam(3) & Format$(aBasicParam(1), "00") & "/", aChkUmaban(aChkTarget), aChkHimo(aChkTarget))
            End If
        End If
    Next ii
    
    Timer1.Enabled = False
End Sub

Private Sub Command48_Click()
    
    'ファイルリスト作成
    ' FileSystemObject (FSO) の新しいインスタンスを生成する
    Dim cFso As FileSystemObject
    Set cFso = New FileSystemObject

    ' Folder オブジェクトを取得する
    Dim cFolder As Folder
    Set cFolder = cFso.GetFolder(App.Path & "\jav\")

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFso = Nothing

    Dim stPrompt As String
    Dim cFile    As file

    ' すべてのファイルを列挙する
    For Each cFile In cFolder.files
        stPrompt = stPrompt & cFile.Path & ","
    Next cFile

    ' 不要になった時点で参照を解放する (Terminate イベントを早めに起こす)
    Set cFolder = Nothing
    Set cFile = Nothing
    
    files = Split(stPrompt, ",")
    
    
    Dim flg1 As Boolean
    Dim licnt As Long
    Dim cnt As Long
    Dim dat() As String
    Dim datcnt As Long
    
    fn = FreeFile
    
    For i = 0 To UBound(files) - 1
        Open files(i) For Input As #fn
        
        Line Input #fn, wk
        
        buns = Split(wk, vbLf)
        licnt = 0
        flg1 = False
        
        For j = 0 To UBound(buns)
            If InStr(buns(j), "<div class=""wp-pagenavi"">") Then
                Exit For
            End If
            If flg1 = True Then
                If licnt > 1 Then
                    ReDim Preserve dat(datcnt)
                    dat(datcnt) = buns(j)
                    datcnt = datcnt + 1
                End If
                If InStr(buns(j), "<li>") Then
                    licnt = licnt + 1
                End If
            Else
                If InStr(buns(j), "<a title=""JavPOP home") Then
                    flg1 = True
                End If
            End If
        Next j
        
        Close #fn
    Next i
    
    fnum2 = FreeFile
    Open App.Path & "\javlist.txt" For Output As #fnum2
    
    For i = 0 To UBound(dat)
        Print #fnum2, dat(i)
    Next i
    
    Close #fnum2
    
    
    
End Sub

