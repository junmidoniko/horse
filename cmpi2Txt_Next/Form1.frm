VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  '�Œ��޲�۸�
   Caption         =   "cmpi 2 text"
   ClientHeight    =   5010
   ClientLeft      =   11460
   ClientTop       =   10035
   ClientWidth     =   9945
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkTF 
      Caption         =   "TargetFrontier"
      Height          =   525
      Left            =   6420
      TabIndex        =   24
      Top             =   2400
      Value           =   2  '����
      Width           =   1785
   End
   Begin VB.Frame fraUser 
      Caption         =   "���[�U�[���"
      Height          =   1395
      Left            =   420
      TabIndex        =   18
      Top             =   3330
      Visible         =   0   'False
      Width           =   9345
      Begin VB.CommandButton Command6 
         Caption         =   "�m��"
         Height          =   885
         Left            =   7170
         TabIndex        =   23
         Top             =   300
         Width           =   1815
      End
      Begin VB.TextBox txtCode 
         Height          =   405
         Left            =   1680
         TabIndex        =   22
         Text            =   "0000"
         Top             =   810
         Width           =   5325
      End
      Begin VB.TextBox txtMail 
         Height          =   405
         Left            =   1680
         TabIndex        =   19
         Text            =   "jun@buhi-buhi.com"
         Top             =   270
         Width           =   5325
      End
      Begin VB.Label Label2 
         Caption         =   "�p�X�R�[�h"
         Height          =   345
         Index           =   1
         Left            =   210
         TabIndex        =   21
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "���[���A�h���X"
         Height          =   345
         Index           =   0
         Left            =   210
         TabIndex        =   20
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��֕���"
      Height          =   585
      Left            =   1620
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��֓����n"
      Height          =   615
      Left            =   2640
      TabIndex        =   16
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optMode 
      Caption         =   "�g��"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.OptionButton optMode 
      Caption         =   "�n��"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   14
      Top             =   1680
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "cnv new"
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtAll 
      Height          =   375
      Left            =   210
      TabIndex        =   12
      Text            =   "c:\test"
      Top             =   120
      Width           =   4215
   End
   Begin VB.CommandButton cmdCmpi 
      Caption         =   "cmpi2file ALL"
      Height          =   795
      Left            =   240
      TabIndex        =   11
      Top             =   660
      Width           =   4185
   End
   Begin VB.ListBox List1 
      Height          =   2040
      Left            =   4560
      OLEDropMode     =   1  '�蓮
      TabIndex        =   10
      Top             =   90
      Width           =   5295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "cnv"
      Height          =   375
      Left            =   4470
      TabIndex        =   9
      Top             =   9270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   6630
      TabIndex        =   7
      Text            =   "C:\test\"
      Top             =   8250
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "exe"
      Height          =   495
      Left            =   4500
      TabIndex        =   6
      Top             =   7980
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   6120
      TabIndex        =   2
      Text            =   "<FONT SIZE=\+2>20.+����"
      Top             =   7500
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   6900
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   6120
      TabIndex        =   0
      Text            =   $"Form1.frx":0000
      Top             =   6300
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.Label Label1 
      Caption         =   "file"
      Height          =   495
      Index           =   3
      Left            =   4500
      TabIndex        =   8
      Top             =   8700
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "pattern"
      Height          =   495
      Index           =   2
      Left            =   4500
      TabIndex        =   5
      Top             =   7380
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "dst"
      Height          =   495
      Index           =   1
      Left            =   4500
      TabIndex        =   4
      Top             =   6780
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "src"
      Height          =   495
      Index           =   0
      Left            =   4380
      TabIndex        =   3
      Top             =   6180
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' �J�����g�f�B���N�g����ύX����API
Private Declare Function SetCurrentDirectory Lib "kernel32" Alias _
    "SetCurrentDirectoryA" (ByVal CurrentDir As String) As Long

Private objNonCode As Object    ' �����R�[�h����/�ϊ��I�u�W�F�N�g
Private strOutCode As String    ' �o�͕����R�[�h

Private Sub txt2Harai(src As String, file As String)
Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String
Dim wkstr As String
Dim dptr As Integer
Dim cutstr As String
Dim tmp As String

Command5.Enabled = False

    
    'HTML�t�@�C��(param.)���������[�ɓW�J
    '<<�t�@�C�� �J>>
    fn = FreeFile
    Open src For Input As #fn
    
    '<<�t�@�C�� ��>>
    lCnt = 0
    Do Until EOF(fn)
        Line Input #fn, wk
        ReDim Preserve data(lCnt)
        data(lCnt) = wk
        lCnt = lCnt + 1
    Loop
    
    '<<�t�@�C�� ��>>
    Close #fn
    
    '<<�f�[�^���>>
    '���K�\���I�u�W�F�N�g�̐錾
    Set objRegExp = New RegExp
    
    With objRegExp
        .Global = True '�����}�b�`��
        .IgnoreCase = True
        .Global = True
        
        '�D�����n ��12��@�D�����n�@��5���@2015�N3��13���V��F���n��F�c�d
        '�J�ÁA�N�������擾
        
        '1 R
        '���[�X�ԍ����擾
        
        '���[�X�ԍ��̂Q�s������A�������擾
        '�P��    ����    �g��    ���ʔn��    �g�P    �n�P
        '�g��    ���z    �l�C    �g��    ���z    �l�C    �g��    ���z    �l�C    �g��    ���z    �l�C    �g��    ���z    �l�C    �g��    ���z    �l�C
        '6   130 1   6   100 1   4-6 220 1   4-6 240 1   6-4 300 1   6-4 360 1
        '-   -   -   4   130 2   -   -   -   -   -   -   -   -   -   -   -   -
        '-   -   -   9   620 8   -   -   -   -   -   -   -   -   -   -   -   -
        
        
        '�Q�s������A�R�A���A�R�A�P���擾
        '���C�h  �O�A��  �O�A�P  ���l
        '�g��    ���z    �l�C    �g��    ���z    �l�C    �g��    ���z    �l�C
        '4-6 150 1   4-6-9   3,200   10  6-4-9   7,020   20
        '6-9 1,080   11  -   -   -   -   -   -
        '4-9 1,580   15  -   -   -   -   -   -
        
        
End With
        
        
        
    '�e�L�X�g�t�@�C��(param.)�֏o��
    src = file
    fn = FreeFile
    Open src For Append As #fn
    
    '<<�t�@�C�� ��>>
    
    wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho
    
    For idx = 1 To 12
        wk2 = ""
        For lCnt = 1 To 20
            wk2 = wk2 & "," & cmpidata(idx, lCnt)
        Next lCnt
        
        wk2 = wk & Format$(idx, "00") & wk2
        Print #fn, wk2
    
    Next idx
    
    '<<�t�@�C�� ��>>
    Close #fn
    Close #wfn
    
    

Command5.Enabled = True
End Sub

Private Sub Command1_Click()
'    <TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
' pattern:<FONT SIZE=\+2>20.+����

'    <TD NOWRAP> �P�q<BR>�T���R��</TD>
' pattern:�q<BR>
'    <TD NOWRAP>�n��<BR>�w��</TD>
' pattern:�n��<BR>�w��
'    <TD BGCOLOR="#FFC0CB" NOWRAP>�Q<BR>84</TD>
' pattern:NOWRAP>.+<
' pattern:<BR>.+<
'    <TD NOWRAP>�X<BR>68</TD>
'    <TD COLSPAN=2 NOWRAP>�@</TD>
' pattern:<TD COLSPAN=2 NOWRAP>


Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim retstr As String '

Command1.Enabled = False

'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    .Pattern = Text1(2).Text    '"[^0-9A-Za-z]"
    
   Set Matches = .Execute(Text1(0).Text)   ' ���������s���܂��B
   For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
      retstr = retstr & "��v���镶���񂪌��������ʒu�́A"
      retstr = retstr & Match.FirstIndex & " �ł��B��v����������́A"
      retstr = retstr & Match.value & " �ł��B" & vbCrLf
   Next
   strResult = retstr
End With

Set objRegExp = Nothing

'���ʂ̕\��
Text1(1).Text = strResult

Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Dim src As String

src = Text1(3).Text

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer

Command2.Enabled = False
Dim wkwk As String
    If optMode(0).value = True Then
         wkwk = "�n��"
    Else
         wkwk = "�g��"
    End If

'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '�J�Ïꏊ�A�N����������
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
        Case 0
             .Pattern = "<FONT SIZE=\+2>20.+����"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "<font size=\+2>20.+����"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = "�N.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "�D�y"
                    basho = "01"
                Case "����"
                    basho = "02"
                Case "����"
                    basho = "03"
                Case "�V��"
                    basho = "04"
                Case "����"
                    basho = "05"
                Case "���R"
                    basho = "06"
                Case "����"
                    basho = "07"
                Case "���s"
                    basho = "08"
                Case "��_"
                    basho = "09"
                Case "���q"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        '���[�X�ԍ�������
        '<TD NOWRAP> �P�q<BR>�T���R��</TD>
        Case 1
            .Pattern = ">.+�q<BR>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+�q<br>"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> MidB$(retstr, 5, 4) Then
                    raceNo = MidB$(retstr, 5, 4)
                    If Right$(raceNo, 1) = "�q" Then
                        raceNo = StrConv(Left$(raceNo, 1), vbNarrow)
                    End If
                    phase = 2
                End If
            End If
        '�R���s�w���f�[�^���O������
        '<TD NOWRAP>�n��<BR>�w��</TD>
        Case 2
             .Pattern = wkwk & "<BR>�w��"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = wkwk & "<br>�w��"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
        
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        '�R���s�w���f�[�^������
        '<TD NOWRAP>�X<BR>68</TD>
        Case 3
             .Pattern = "NOWRAP>.+<"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "nowrap>.+<"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos = 0 Then
                '�I�[�`�F�b�N
                '<TD COLSPAN=2 NOWRAP>�@</TD>
                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "<td colspan=. nowrap>"
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
            
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data ��荞��
                'NOWRAP>11<BR>71<
                 
                 'umaban
                 .Pattern = "NOWRAP>.+<BR>"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap>.+<br>"
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If Mid$(wk, 9, 1) = "<" Then
                    wk = Mid$(wk, 8, 1)
                Else
                    wk = Mid$(wk, 8, 2)
                End If
                
                If IsNumeric(wk) = False Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                Else
                    umaban = wk
                    
                    cmpidata(CInt(raceNo), umaban) = cmpininki
                    
                     'value
                     .Pattern = "<BR>.+<"
                    pos = 0
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 5, 1)
                    Else
                        value = Mid$(wk, 5, 2)
                    End If
                    cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                    
                    retstr = retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'���ʂ̕\��
Text1(1).Text = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho


'�e�L�X�g�t�@�C��(param.)�֏o��
src = Text1(3).Text & ".txt"
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

For idx = 1 To 12
    wk2 = ""
    For lCnt = 1 To 20
        wk2 = wk2 & "," & cmpidata(idx, lCnt)
    Next lCnt
    
    wk2 = wk & Format$(idx, "00") & wk2
    Print #fn, wk2

Next idx

'<<�t�@�C�� ��>>
Close #fn

Command2.Enabled = True

End Sub

Private Sub Command3_Click()
Dim src As String

src = Text1(3).Text

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer

Command3.Enabled = False
Dim wkwk As String
    If optMode(0).value = True Then
         wkwk = "�n��"
    Else
         wkwk = "�g��"
    End If

'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '�J�Ïꏊ�A�N����������
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
        Case 0
             .Pattern = "<FONT SIZE=\+2>20.+����"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "<font size=\+2>20.+����"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = "�N.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "�D�y"
                    basho = "01"
                Case "����"
                    basho = "02"
                Case "����"
                    basho = "03"
                Case "�V��"
                    basho = "04"
                Case "����"
                    basho = "05"
                Case "���R"
                    basho = "06"
                Case "����"
                    basho = "07"
                Case "���s"
                    basho = "08"
                Case "��_"
                    basho = "09"
                Case "���q"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        '���[�X�ԍ�������
        '<TD NOWRAP> �P�q<BR>�T���R��</TD>
        Case 1
            .Pattern = ">.+�q<BR>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+�q<br>"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> MidB$(retstr, 5, 4) Then
                    raceNo = MidB$(retstr, 5, 4)
                    If Right$(raceNo, 1) = "�q" Then
                        raceNo = StrConv(Left$(raceNo, 1), vbNarrow)
                    End If
                    phase = 2
                End If
            End If
        '�R���s�w���f�[�^���O������
        '<TD NOWRAP>�n��<BR>�w��</TD>
        Case 2
             .Pattern = wkwk & "<BR>�w��"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = wkwk & "<br>�w��"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
        
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        '�R���s�w���f�[�^������
        '<TD NOWRAP>�X<BR>68</TD>
        Case 3
             .Pattern = "NOWRAP>.+<"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "nowrap>.+<"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos = 0 Then
                '�I�[�`�F�b�N
                '<TD COLSPAN=2 NOWRAP>�@</TD>
                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "<td colspan=. nowrap>"
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
            
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data ��荞��
                'NOWRAP>11<BR>71<
                 
                 'umaban
                 .Pattern = "NOWRAP>.+<BR>"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap>.+<br>"
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If Mid$(wk, 9, 1) = "<" Then
                    wk = Mid$(wk, 8, 1)
                Else
                    wk = Mid$(wk, 8, 2)
                End If
                
                If IsNumeric(wk) = False Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                Else
                    umaban = wk
                    
                    cmpidata(CInt(raceNo), umaban) = cmpininki
                    
                     'value
                     .Pattern = "<BR>.+<"
                    pos = 0
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 5, 1)
                    Else
                        value = Mid$(wk, 5, 2)
                    End If
                    cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                    
                    retstr = retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'���ʂ̕\��
Text1(1).Text = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho


'�e�L�X�g�t�@�C��(param.)�֏o��
src = Text1(3).Text & ".txt"
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

For idx = 1 To 12
    wk2 = ""
    For lCnt = 1 To 20
        wk2 = wk2 & "," & cmpidata(idx, lCnt)
    Next lCnt
    
    wk2 = wk & Format$(idx, "00") & wk2
    Print #fn, wk2

Next idx

'<<�t�@�C�� ��>>
Close #fn

Command3.Enabled = True
End Sub


Private Sub Command4_Click()
    Dim i As Integer
    Dim src As String
    Dim file As String
    Dim wfile As String
    
    file = txtAll.Text & "\cmpi" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    wfile = txtAll.Text & "\cmpN" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
'        Call msHTML2Txt(src, file)
        Call nankan2Txt(src & ".txt", file, wfile)
    Next i

End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Dim src As String
    Dim file As String
    
    file = txtAll.Text & "\harai" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
        Call txt2Harai(src & ".txt", file)
    Next i

End Sub

Private Sub Command6_Click()
    '���[���A�h���X���Í���
    
    Dim key() As Byte
    Dim iv() As Byte
    Dim data() As Byte
    Dim objCipher As Cipher
    Dim retDat As String
    Dim retdata As String
    
    key = StringUtility.stringToByte("27842784midoniko")
    iv = StringUtility.stringToByte("midoniko27842784")
    data = StringUtility.stringToByte(txtMail.Text)

    On Error GoTo ErrorHandler
    Set objCipher = New Cipher

    Call objCipher.encrypt(key, iv, data)
    retdata = Base64.encode(data)
    
    Dim i As Integer
    Dim hit As Boolean
    hit = False
    
'    For i = 0 To UBound(gMail)
'        If gMail(i) = retdata Then
'            hit = True
'            optMode(0).Visible = True
'            Command4.Visible = True
'            fraUser.Visible = False
'        End If
'    Next i
    
    Exit Sub

ErrorHandler:
    Dim message As String

    message = "�G���[�R�[�h: &H" & Hex(Err.Number) & vbCrLf & _
        "�\�[�X: " & Err.Source & vbCrLf & Err.Description
        MsgBox message, vbCritical
End Sub

Private Sub Form_Load()
    
    ReDim gMail(22)
    
    '�Í������ꂽ���[���A�h���X C:\temp\ango apl�ɔz�u
    gMail(0) = "fqmYD2TydYK5aSARSuLbmt8PgC2BQBeVUpO3bmFeQKg="       'jun@
    gMail(1) = "AnS/VZJd80Anj28N6+nYx0E/Z0NNxZ9I+gX0ctDe+tQ="
    gMail(2) = "3uOGzpZTJyIruZt8OA6WcituQwv5MP3fafWFrZCL/M0="
    gMail(3) = "1eK4gXvPBkWXR34FGh+6OhbHfKFwrCUQzLq23QzwDHs="
    gMail(4) = "Zkq0J1OEZqvfi30hi+qlxAeeFjP+y4IP9d5kxhwNKLg="
    gMail(5) = "1eK4gXvPBkWXR34FGh+6OhbHfKFwrCUQzLq23QzwDHs="
    gMail(6) = "KIBdlZwrNPZ+6s8tVLzCicQSCpRlDOKxuuRMoJU78X0="
    gMail(7) = "HnDoMAl1zACZK94Stx/+bopUsGjRQYN1SndyIs4yzuI="
    gMail(8) = "5ZvYFJdk9VlACkBYlMCIQhl4lQOELhB7gUgiPn2KY+c="
    gMail(9) = "hzC1KuAEzdk1Ua2TjBoxRbZFkdE9vzMfQRTRCEO8jyo="
    gMail(10) = "KXFlLAafK5ugR+TabFtaxq+31uiPmhqVl/qs7iqJ2ds="
    gMail(11) = "uQM0k7fJyp+Zz/1KNBpFxChZCx97eb9P8bsKpLcbdEw="
    gMail(12) = "eb8p8f4H7vTIbS/QiV8CGh6nJwVttzdJC8zp06J5R38="
    gMail(13) = "LZKwOSEcUsb4DMFXqDyh30GvoAV5DAiDqjKQeDEJwb8="
    gMail(14) = "uA47JH8Vsscz7s/he3xE/ZC6m9GalfvF5Gx1VodoCPg="
    gMail(15) = "1YfrSfN1dDrTF7KjSkCBMgf/kHjlLuwGzBOJmeXl6n4="
    gMail(16) = "3M0YnwW9JdMCKymYLZi8sqM8lxyWLrOy3pWgQ/m2kFegTLkjn0KSW6cZ92mSD6z4"
    gMail(17) = "AS4tjAzHF75HLqlyvJreoveVz1eNWIg7gEHRsgb6O+k="
    gMail(18) = "p0qtX1hbCzNeCIWRi8IdydY3EykCsmRDcqXEgEMqBcA="
    gMail(19) = "sVEcUkpxh8ZiWE1GbUWElzDy4wzaJmdBTKaOpZTSwkg="
    gMail(20) = "TdWtVUJ740Hhbq08AjZ/uvlb6+AwmjJqr0NfqFn9tRY="
    gMail(21) = "8SyijqDiFbMmQczsmu4OKSFjITA+5Ht3rnnT2744jYI="
    gMail(22) = "E8yEkVqG0ZCcm9y4Jbwm1ATRlqtcv+peO8C1giSfpQU="
    
    '�g�p�����`�F�b�N
    Dim nowD As String
    nowD = Format$(Now, "yyyymmdd")
'    If nowD > "20170705" Then                   '20170403
'        MsgBox "�g�p�������؂�܂����B�z�[���y�[�W����ŐV�ł��_�E�����[�h���Ă��������B"
'        End
'    End If
    
    
    ' �J�����g�t�H���_�̕ύX
    SetCurrentDirectory App.Path
    
    ' NonCodeVb6.dll�̑��݃`�F�b�N
    On Error Resume Next
    If CreateObject("NonCodeVb6.NonCodeClass") Is Nothing Then
        If Len(Dir("NonCodeVb6.dll")) <> 0 Then
            ' NonCodeVb6.dll�̃��W�X�g���o�^
            Shell "regsvr32 /s NonCodeVb6.dll", vbHide
        Else
            ' NonCodeVb6.dll��Code2Code.exe�Ɠ����t�H���_�ɒu���Ă��������B
            MsgBox _
            "NonCodeVb6.dll��������܂���ł����B" & vbCrLf & vbCrLf & _
            "NonCodeVb6.dll��" & vbCrLf & "[" & App.Path & "]" & vbCrLf & _
            "�ɒu���Ă��������B"
            End
        End If
    End If
    On Error GoTo 0
    
    Set objNonCode = CreateObject("NonCodeVb6.NonCodeClass")
    
    Dim aleft As String
    Dim atop As String
    Dim file As String
    
    gRet = loadIni(App.Title, G_INI_SEC_WINDOW, G_INI_KEY_WINDOW_TOP, atop)
    gRet = loadIni(App.Title, G_INI_SEC_WINDOW, G_INI_KEY_WINDOW_LEFT, aleft)
    gRet = loadIni(G_COMMON_INIFILE, G_INI_SEC_CMPI, G_INI_KEY_CMPI_TXT, file)
    '�R���s�f�[�^�o�̓t�@�C��
    file = "C:\test" 'App.Path & "\" & file
    
    Top = 0 'CInt(atop)
    Left = 0 'CInt(aleft)
    txtAll.Text = file
    
    ' SJIS�ւ̕ϊ�
    strOutCode = "SJIS"

'    MsgBox "�o�̓t�@�C���̃f�[�^���폜���܂����H"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim aleft As Integer
    Dim atop As Integer
    
    aleft = Left
    atop = Top
    
    gRet = saveIni(App.Title, G_INI_SEC_WINDOW, G_INI_KEY_WINDOW_TOP, CStr(atop))
    gRet = saveIni(App.Title, G_INI_SEC_WINDOW, G_INI_KEY_WINDOW_LEFT, CStr(aleft))
End Sub

Private Sub List1_DblClick()
    List1.Clear
End Sub

Private Sub List1_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lstrTmp             As String
    Dim i As Integer
    
On Error GoTo ErrHandler
    
    '��ۯ�߂��ꂽ���̂��A̧�قł��邩���f
    If data.GetFormat(vbCFFiles) Then
        For i = 1 To data.Files.Count
            List1.AddItem (data.Files(i))
        Next i
        
    Else
        MsgBox "�h���b�v���ꂽ���̂�̧�قł͂���܂���B"
        Exit Sub
    End If
    
    Exit Sub
ErrHandler:
    MsgBox "error:" & Err.Description
    Exit Sub
End Sub

'���X�g�ɂ���Cmpi�f�[�^�iHTML�j���e�L�X�g�t�@�C���Ɉꊇ�ϊ�
Private Sub cmdCmpi_Click()
    Dim i As Integer
    Dim src As String
    Dim file As String
    Dim wfile As String
    
    file = txtAll.Text & "\cmpi" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    wfile = txtAll.Text & "\cmpW" & Format$(Now, "yyyymmddhhnnss") & ".txt" 'txtAll.Text
    
    For i = 0 To List1.ListCount - 1
        src = List1.List(i)
        Call TextCodeChg(src)
'        Call msHTML2Txt(src, file)
        Call Compi2Txt(src & ".txt", file, wfile)
    Next i
    
End Sub

Private Sub msHTML2Txt(src As String, file As String)

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String

cmdCmpi.Enabled = False
Dim wkwk As String
    If optMode(0).value = True Then
         wkwk = "�n��"
    Else
         wkwk = "�g��"
    End If

'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '�J�Ïꏊ�A�N����������
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
        Case 0
             .Pattern = "<FONT SIZE=\+2>20.+����"
'             .Pattern = "<font size=""\+2"">20.+����"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "<font size=""\+2"">20.+����"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = "�N.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+��.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 6 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "�D�y"
                    basho = "01"
                Case "����"
                    basho = "02"
                Case "����"
                    basho = "03"
                Case "�V��"
                    basho = "04"
                Case "����"
                    basho = "05"
                Case "���R"
                    basho = "06"
                Case "����"
                    basho = "07"
                Case "���s"
                    basho = "08"
                Case "��_"
                    basho = "09"
                Case "���q"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        '���[�X�ԍ�������
        '<TD NOWRAP> �P�q<BR>�T���R��</TD>
        Case 1
            .Pattern = ">.+�q<BR>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+�q<br>"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> MidB$(retstr, 5, 4) Then
                    wkRaceNo = MidB$(retstr, 5, 4)
                    If Right$(wkRaceNo, 1) = "�q" Then
                        wkRaceNo = StrConv(Left$(wkRaceNo, 1), vbNarrow)
                    End If
                    If wkRaceNo <> raceNo And IsNumeric(wkRaceNo) Then
                        raceNo = wkRaceNo
                        phase = 2
                    End If
                End If
            End If
        '�R���s�w���f�[�^���O������
        '<TD NOWRAP>�n��<BR>�w��</TD>
        Case 2
             .Pattern = wkwk & "<BR>�w��"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = wkwk & "<br>�w��"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        '�R���s�w���f�[�^������
        '<TD NOWRAP>�X<BR>68</TD>
        Case 3
             .Pattern = "NOWRAP>.+<"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = "nowrap>.+<"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos = 0 Then
                '�I�[�`�F�b�N
                '<TD COLSPAN=2 NOWRAP>�@</TD>
                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "<td colspan=. nowrap>"
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
                
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data ��荞��
                'NOWRAP>11<BR>71<
                 
                 'umaban
                 .Pattern = "NOWRAP>.+<BR>"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap>.+<br>"
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If Mid$(wk, 9, 1) = "<" Then
                    wk = Mid$(wk, 8, 1)
                Else
                    wk = Mid$(wk, 8, 2)
                End If
                
                If IsNumeric(wk) = False Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                Else
                    umaban = wk
                    
'                    backup = cmpidata(CInt(raceNo), umaban)
                    
                     'value
                     .Pattern = "<BR>.+<"
                    pos = 0
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 5, 1)
                    Else
                        value = Mid$(wk, 5, 2)
                    End If
                    If IsNumeric(value) = False And value <> "��" Then
                        If raceNo = "12" Then
                            Exit For
                        Else
                            phase = 1
                        End If
                    Else
                        cmpidata(CInt(raceNo), umaban) = cmpininki
                        cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                    End If
                    
                    retstr = retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'�e�L�X�g�t�@�C��(param.)�֏o��
src = file
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

For idx = 1 To 12
    wk2 = ""
    For lCnt = 1 To 20
        wk2 = wk2 & "," & cmpidata(idx, lCnt)
    Next lCnt
    
    wk2 = wk & Format$(idx, "00") & wk2
    Print #fn, wk2

Next idx

'<<�t�@�C�� ��>>
Close #fn

cmdCmpi.Enabled = True

End Sub

Private Sub Compi2Txt(src As String, file As String, wfile As String)

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fnTfr As Integer
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wk3 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim cmpiTfr(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String

cmdCmpi.Enabled = False

Dim wkwk As String
    If optMode(0).value = True Then
         wkwk = "�n��"
    Else
         wkwk = "wakuNum"
    End If


'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

'wfn = FreeFile
'Open wfile For Append As #wfn


'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '�J�Ïꏊ�A�N����������
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
        '<h2 id="contentTit">2012�N1��5���@�R���s�w���|1�񒆎R1����</h2>
        Case 0
'             .Pattern = "<FONT SIZE=\+2>20.+����"
             .Pattern = "<h2 id=""contentTit"">.+����"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
'                .Pattern = "<font size=""\+2"">20.+����"
                .Pattern = "contentTit""\>20.+����"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = "�N.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+���@�R���s�w��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 9 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "�D�y"
                    basho = "01"
                Case "����"
                    basho = "02"
                Case "����"
                    basho = "03"
                Case "�V��"
                    basho = "04"
                Case "����"
                    basho = "05"
                Case "���R"
                    basho = "06"
                Case "����"
                    basho = "07"
                Case "���s"
                    basho = "08"
                Case "��_"
                    basho = "09"
                Case "���q"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        '���[�X�ԍ�������
        '<TD NOWRAP> �P�q<BR>�T���R��</TD>
        '<td class="racename"><span class="race">12R</span>
        Case 1
'            .Pattern = ">.+�q<BR>"
            .Pattern = ">.+R\<\/span\>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+�q<br>"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> Mid$(retstr, 21, 2) Then
                    wkRaceNo = Mid$(retstr, 21, 2)
                    If Right$(wkRaceNo, 1) = "R" Then
                        wkRaceNo = StrConv(Left$(wkRaceNo, 1), vbNarrow)
                    End If
                    If wkRaceNo <> raceNo And IsNumeric(wkRaceNo) Then
                        raceNo = wkRaceNo
                        phase = 2
                    End If
                End If
            End If
        '�R���s�w���f�[�^���O������
        '<TD NOWRAP>�n��<BR>�w��</TD>
        Case 2
'             .Pattern = "�n��<BR>�w��"
            If optMode(0).value = True Then
                .Pattern = "�n��" & "\<br\>�w��" '0504"\<br \/\>�w��"
            Else
                .Pattern = "wakuNum.+\>8"
            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                If optMode(0).value = True Then
                    .Pattern = wkwk & "<br>�w��"
                Else
                    .Pattern = "wakuNum.+\>8"
                End If
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        '�R���s�w���f�[�^������
        '<TD NOWRAP>�X<BR>68</TD>
        Case 3
'             .Pattern = "NOWRAP>.+<"
            If optMode(0).value = True Then         '0524
                .Pattern = "\>.+\<br\>.+\<\/td\>"
            Else
                
                ReDim wakD(0)
                
                Do
                    'wakuren�ȍ~��</td>�ŋ�؂�
                    plc = InStr(data(lCnt), "</td>")
                    If plc > 0 Then
                        smpl = data(lCnt)
                        Do
                            ReDim Preserve wakD(UBound(wakD) + 1)
                            kire = Left$(smpl, plc + 4)
                            wakD(UBound(wakD)) = kire
                            kire = Mid$(smpl, plc + 5)
                            plc = InStr(kire, "</td>")
                            If plc > 0 Then
                                smpl = kire
                            Else
                                Exit Do
                            End If
                        Loop
                    Else
                        '�s�v�f�[�^
                    End If
                    
                    If lCnt = UBound(data) Then
                        Exit Do
                    End If
                    
                    lCnt = lCnt + 1
                Loop
                
                lCnt = 0
                wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho
                wkPrt = wk
                Do
                    If UBound(wakD) < lCnt Then
'                        Print #wfn, wkPrt
                        Exit Do
                    End If
                    If wakD(lCnt) <> "" Then
                        retstr = ""
                        '���[�X�ԍ��H
                        .Pattern = "race""\>.+R\<\/span\>"
                        pos = 0
                        Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           retstr = Match.value
                        Next
                        
                        If retstr <> "" Then
                            '���[�X�ԍ����o
                            kire = Mid$(retstr, 7, 2)
                            If Right$(kire, 1) = "R" Then
                                raceNo = "0" & Left$(kire, 1)
                            Else
                                raceNo = kire
                            End If
                            
                            If wkPrt <> wk Then
'                                Print #wfn, wkPrt
                                wkPrt = wk
                            End If
                            
                            wkPrt = wkPrt & "," & raceNo
                        Else
                            '�g�Ԓ��o
                            .Pattern = "\>.+\<br \/\>"
                            pos = 0
                            Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                               retstr = Match.value
                            Next
                            
                            wban = Mid$(retstr, 2, 2)
                            If wban <> "�g��" And wban <> "" Then
                            
                                wkPrt = wkPrt & "," & wban
                                
                                '�R���s�w�����o
                                .Pattern = "\>.+?\<\/td\>"
                                pos = 0
                                Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                                   retstr = Match.value
                                Next
                                
                                cmpV = Left$(Right$(retstr, 7), 2)
                                If cmpV = "���" Then
                                    cmpV = "00"
                                End If
                                wkPrt = wkPrt & "," & cmpV
                            End If
                        End If
                    End If
                    
                    lCnt = lCnt + 1
                Loop
                
'                Close #wfn
                
                cmdCmpi.Enabled = True
                
                Exit Sub
            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                If optMode(0).value = True Then
                    .Pattern = "nowrap>.+<"
                Else
                    .Pattern = "�g��.+\>.+\<br \/\>.+\<\/td\>"
                End If
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            
            If pos = 0 Then
                '�I�[�`�F�b�N
                '<TD COLSPAN=2 NOWRAP>�@</TD>
'                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 .Pattern = "�@\<\/td\>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "tr\>"
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
                
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data ��荞��
                'NOWRAP>11<BR>71<
                ' 6<br />86</td>
                 
                 'umaban
                '.Pattern = "NOWRAP>.+<BR>"
                If InStr(retstr, "�g��") = 0 Then
                    .Pattern = "\>.+\<br\>.+\<\/td\>"
                Else
                    .Pattern = "race""\>.+R\<"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                    
                    raceNo = Mid$(retstr, 7, 2)
                    If Right$(raceNo, 1) = "R" Then
                        raceNo = Left$(raceNo, 1)
                    End If
                    
                    .Pattern = "�g��.+\<br \/\>.+\<\/td\>"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
                
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap\>.+\<br\>"
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If InStr(retstr, "suishou") = 0 Then
                    If Mid$(wk, 9, 1) = "<" Then
                        wk = Mid$(wk, 2, 2)
                    Else
                        wk = Mid$(wk, 2, 2)
                    End If
                Else
                    wk = Mid$(wk, 36, 2)
                End If
                
                If optMode(1).value = True Then
                    wakCnt = wakCnt + 1
                End If
                
                If IsNumeric(wk) = False Or wakCnt = 8 Then
                    wakCnt = 0
                    
                    If raceNo = "12" Then
                        Exit For
                    Else
                        If optMode(1).value = True Then
                            phase = 3
                        Else
                            phase = 1
                        End If
                    End If
                Else
                    umaban = wk
                    
'                    backup = cmpidata(CInt(raceNo), umaban)
                    
                     'value
                     .Pattern = "\<br\>.+\<"        '0524
                    pos = 0
                    
                    If optMode(1).value = True Then
                        retstr = Right$(retstr, 13)
                    End If
                    
                    
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 5, 2)
                    Else
                        value = Mid$(wk, 5, 2)
                    End If
                    If IsNumeric(value) = False And Left$(value, 1) <> "��" Then
                        If raceNo = "12" Then
                            Exit For
                        Else
                            phase = 1
                        End If
                    Else
                        '������������P�[�X��Ή� 20170524
                        If Left$(value, 1) = "��" Then
                            value = "0"
                        End If
                        cmpidata(CInt(raceNo), umaban) = cmpininki
                        cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                        cmpiTfr(CInt(raceNo), umaban) = value
                        
                        Do
                            If InStr(Mid$(retstr, InStr(retstr, "��") + 6), "��") = 0 Then
                                Exit Do
                            End If
                            
                            cmpininki = cmpininki + 1
                            retstr = Mid$(retstr, InStr(retstr, "��") + 6)
                            
                            .Pattern = "\<td\>.+\<"        '0524
                            
                            Set Matches = .Execute(retstr)   ' ���������s���܂��B
                            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                               wk = Match.value
                            Next
                            
                            If Mid$(wk, 6, 1) = "<" Then
                                value = Mid$(wk, 5, 2)
                            Else
                                value = Mid$(wk, 5, 2)
                            End If
                            
                            umaban = CInt(value)
                            cmpidata(CInt(raceNo), umaban) = cmpininki
                            cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & "0"
                            cmpiTfr(CInt(raceNo), umaban) = "0"
                        Loop
                        
                    End If
                    
                    retstr = "" 'retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'�e�L�X�g�t�@�C��(param.)�֏o��
src = file
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

Dim tfr As String


    For idx = 1 To 12
        
        tfr = txtAll.Text & "\" & wk & Format$(idx, "00") & ".csv"
        fnTfr = FreeFile
If chkTF.value = 1 Then
        Open tfr For Output As #fnTfr
End If
        
        wk2 = ""
        For lCnt = 1 To 20
            wk2 = wk2 & "," & cmpidata(idx, lCnt)
            If cmpiTfr(idx, lCnt) <> "" Then
    '            If wk3 = "" Then
    '                wk3 = cmpiTfr(idx, lCnt)
    '            Else
                    wk3 = wk3 & "," & cmpiTfr(idx, lCnt)
    '            End If
            End If
        Next lCnt
        
        wk2 = wk & Format$(idx, "00") & wk2
        Print #fn, wk2
If chkTF.value = 1 Then
        Print #fnTfr, wk & Format$(idx, "00") & wk3
        
        Close #fnTfr
End If
    
    Next idx

'<<�t�@�C�� ��>>
Close #fn
'Close #wfn

cmdCmpi.Enabled = True

End Sub


Private Sub Compi2Txt_res(src As String, file As String, wfile As String)

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fnTfr As Integer
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wk3 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim cmpiTfr(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String

cmdCmpi.Enabled = False

Dim wkwk As String
'    If optMode(0).value = True Then
'         wkwk = "�n��"
'    Else
'         wkwk = "wakuNum"
'    End If


'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

wfn = FreeFile
Open wfile For Append As #wfn


'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '"��@��苣�n�@��1���@<span class="tx-small tx-normal">2015�N4��1��</span>"������
        '��1��@��苣�n�@��1���@<span class="tx-small tx-normal">2015�N4��1��</span></div>
        Case 0
'             .Pattern = "<FONT SIZE=\+2>20.+����"
             .Pattern = "��@.+���n"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
'            If pos = 0 Then
''                .Pattern = "<font size=""\+2"">20.+����"
'                .Pattern = "="".+R"
'                Set Matches = .Execute(Data(lCnt))   ' ���������s���܂��B
'                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
'                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
'                   retstr = Match.value
'                Next
'            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = """.+R"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 2 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+���@�R���s�w��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 9 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "�D�y"
                    basho = "01"
                Case "����"
                    basho = "02"
                Case "����"
                    basho = "03"
                Case "�V��"
                    basho = "04"
                Case "����"
                    basho = "05"
                Case "���R"
                    basho = "06"
                Case "����"
                    basho = "07"
                Case "���s"
                    basho = "08"
                Case "��_"
                    basho = "09"
                Case "���q"
                    basho = "10"
                End Select
                
                phase = 1
            End If
        '���[�X�ԍ�������
        '<TD NOWRAP> �P�q<BR>�T���R��</TD>
        '<td class="racename"><span class="race">12R</span>
        Case 1
'            .Pattern = ">.+�q<BR>"
            .Pattern = ">.+R\<\/span\>"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                .Pattern = ">.+�q<br>"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                If raceNo <> Mid$(retstr, 21, 2) Then
                    wkRaceNo = Mid$(retstr, 21, 2)
                    If Right$(wkRaceNo, 1) = "R" Then
                        wkRaceNo = StrConv(Left$(wkRaceNo, 1), vbNarrow)
                    End If
                    If wkRaceNo <> raceNo And IsNumeric(wkRaceNo) Then
                        raceNo = wkRaceNo
                        phase = 2
                    End If
                End If
            End If
        '�R���s�w���f�[�^���O������
        '<TD NOWRAP>�n��<BR>�w��</TD>
        Case 2
'             .Pattern = "�n��<BR>�w��"
            If optMode(0).value = True Then
                .Pattern = "�n��" & "\<br \/\>�w��"
            Else
                .Pattern = "wakuNum.+\>8"
            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                If optMode(0).value = True Then
                    .Pattern = wkwk & "<br />�w��"
                Else
                    .Pattern = "wakuNum.+\>8"
                End If
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                phase = 3
                cmpininki = 0
            End If
        
        '�R���s�w���f�[�^������
        '<TD NOWRAP>�X<BR>68</TD>
        Case 3
'             .Pattern = "NOWRAP>.+<"
            If optMode(0).value = True Then
                .Pattern = "\>.+\<br \/\>.+\<\/td\>"
            Else
                
                ReDim wakD(0)
                
                Do
                    'wakuren�ȍ~��</td>�ŋ�؂�
                    plc = InStr(data(lCnt), "</td>")
                    If plc > 0 Then
                        smpl = data(lCnt)
                        Do
                            ReDim Preserve wakD(UBound(wakD) + 1)
                            kire = Left$(smpl, plc + 4)
                            wakD(UBound(wakD)) = kire
                            kire = Mid$(smpl, plc + 5)
                            plc = InStr(kire, "</td>")
                            If plc > 0 Then
                                smpl = kire
                            Else
                                Exit Do
                            End If
                        Loop
                    Else
                        '�s�v�f�[�^
                    End If
                    
                    If lCnt = UBound(data) Then
                        Exit Do
                    End If
                    
                    lCnt = lCnt + 1
                Loop
                
                lCnt = 0
                wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho
                wkPrt = wk
                Do
                    If UBound(wakD) < lCnt Then
                        Print #wfn, wkPrt
                        Exit Do
                    End If
                    If wakD(lCnt) <> "" Then
                        retstr = ""
                        '���[�X�ԍ��H
                        .Pattern = "race""\>.+R\<\/span\>"
                        pos = 0
                        Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           retstr = Match.value
                        Next
                        
                        If retstr <> "" Then
                            '���[�X�ԍ����o
                            kire = Mid$(retstr, 7, 2)
                            If Right$(kire, 1) = "R" Then
                                raceNo = "0" & Left$(kire, 1)
                            Else
                                raceNo = kire
                            End If
                            
                            If wkPrt <> wk Then
                                Print #wfn, wkPrt
                                wkPrt = wk
                            End If
                            
                            wkPrt = wkPrt & "," & raceNo
                        Else
                            '�g�Ԓ��o
                            .Pattern = "\>.+\<br \/\>"
                            pos = 0
                            Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                               retstr = Match.value
                            Next
                            
                            wban = Mid$(retstr, 2, 2)
                            If wban <> "�g��" And wban <> "" Then
                            
                                wkPrt = wkPrt & "," & wban
                                
                                '�R���s�w�����o
                                .Pattern = "\>.+?\<\/td\>"
                                pos = 0
                                Set Matches = .Execute(wakD(lCnt))   ' ���������s���܂��B
                                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                                   retstr = Match.value
                                Next
                                
                                cmpV = Left$(Right$(retstr, 7), 2)
                                If cmpV = "���" Then
                                    cmpV = "00"
                                End If
                                wkPrt = wkPrt & "," & cmpV
                            End If
                        End If
                    End If
                    
                    lCnt = lCnt + 1
                Loop
                
                Close #wfn
                
                cmdCmpi.Enabled = True
                
                Exit Sub
            End If
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
                If optMode(0).value = True Then
                    .Pattern = "nowrap>.+<"
                Else
                    .Pattern = "�g��.+\>.+\<br \/\>.+\<\/td\>"
                End If
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            
            If pos = 0 Then
                '�I�[�`�F�b�N
                '<TD COLSPAN=2 NOWRAP>�@</TD>
'                 .Pattern = "<TD COLSPAN=. NOWRAP>"
                 .Pattern = "�@\<\/td\>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "tr\>"
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
                
                If pos <> 0 Then
                    If raceNo = "12" Then
                        Exit For
                    Else
                        phase = 1
                    End If
                End If
            Else
                'data ��荞��
                'NOWRAP>11<BR>71<
                ' 6<br />86</td>
                 
                 'umaban
                '.Pattern = "NOWRAP>.+<BR>"
                If InStr(retstr, "�g��") = 0 Then
                    .Pattern = "\>.+\<br \/\>.+\<\/td\>"
                Else
                    .Pattern = "race""\>.+R\<"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                    
                    raceNo = Mid$(retstr, 7, 2)
                    If Right$(raceNo, 1) = "R" Then
                        raceNo = Left$(raceNo, 1)
                    End If
                    
                    .Pattern = "�g��.+\<br \/\>.+\<\/td\>"
                    pos = 0
                    Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       retstr = Match.value
                    Next
                End If
                
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If pos = 0 Then
                    .Pattern = "nowrap\>.+\<br\>"
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                End If
                
                cmpininki = cmpininki + 1
                If InStr(retstr, "suishou") = 0 Then
                    If Mid$(wk, 9, 1) = "<" Then
                        wk = Mid$(wk, 2, 2)
                    Else
                        wk = Mid$(wk, 2, 2)
                    End If
                Else
                    wk = Mid$(wk, 36, 2)
                End If
                
                If optMode(1).value = True Then
                    wakCnt = wakCnt + 1
                End If
                
                If IsNumeric(wk) = False Or wakCnt = 8 Then
                    wakCnt = 0
                    
                    If raceNo = "12" Then
                        Exit For
                    Else
                        If optMode(1).value = True Then
                            phase = 3
                        Else
                            phase = 1
                        End If
                    End If
                Else
                    umaban = wk
                    
'                    backup = cmpidata(CInt(raceNo), umaban)
                    
                     'value
                     .Pattern = "\<br \/\>.+\<"
                    pos = 0
                    
                    If optMode(1).value = True Then
                        retstr = Right$(retstr, 13)
                    End If
                    
                    
                    Set Matches = .Execute(retstr)   ' ���������s���܂��B
                    For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                       pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                       wk = Match.value
                    Next
                    If pos = 0 Then
                        .Pattern = "<br>.+<"
                        Set Matches = .Execute(retstr)   ' ���������s���܂��B
                        For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                           pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                           wk = Match.value
                        Next
                    End If
                    
                    If Mid$(wk, 6, 1) = "<" Then
                        value = Mid$(wk, 7, 2)
                    Else
                        value = Mid$(wk, 7, 2)
                    End If
                    If IsNumeric(value) = False And Left$(value, 1) <> "��" Then
                        If raceNo = "12" Then
                            Exit For
                        Else
                            phase = 1
                        End If
                    Else
                        If Left$(value, 1) = "��" Then
                            value = "0"
                        End If
                        cmpidata(CInt(raceNo), umaban) = cmpininki
                        cmpidata(CInt(raceNo), umaban) = cmpidata(CInt(raceNo), umaban) & "," & value
                        cmpiTfr(CInt(raceNo), umaban) = value
                    End If
                    
                    retstr = "" 'retstr
                End If
            End If
        
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'�e�L�X�g�t�@�C��(param.)�֏o��
src = file
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

Dim tfr As String
For idx = 1 To 12
    
    tfr = txtAll.Text & "\" & wk & Format$(idx, "00") & ".csv"
    fnTfr = FreeFile
    Open tfr For Output As #fnTfr
    
    wk2 = ""
    For lCnt = 1 To 20
        wk2 = wk2 & "," & cmpidata(idx, lCnt)
        If cmpiTfr(idx, lCnt) <> "" Then
'            If wk3 = "" Then
'                wk3 = cmpiTfr(idx, lCnt)
'            Else
                wk3 = wk3 & "," & cmpiTfr(idx, lCnt)
'            End If
        End If
    Next lCnt
    
    wk2 = wk & Format$(idx, "00") & wk2
    Print #fn, wk2
    Print #fnTfr, wk & Format$(idx, "00") & wk3
    
    Close #fnTfr

Next idx

'<<�t�@�C�� ��>>
Close #fn
Close #wfn

cmdCmpi.Enabled = True

End Sub

Private Sub TextCodeChg(pSrc As String)
    ' �e�L�X�g���o�C�g�z��œǍ�
    Dim ipath As String: ipath = pSrc   'App.Path & "\TestUtf8.txt"
    Dim idat() As Byte
    ReDim idat(FileLen(ipath) - 1) As Byte
    Dim intFileNo As Integer
    intFileNo = FreeFile
    Open ipath For Binary As intFileNo
    Get intFileNo, , idat
    Close intFileNo
            
    ' �����R�[�h����(blnBin=�o�C�i�����薳��)
    Dim cod As String: cod = objNonCode.GetCodeName(idat, blnBin:=False)

    ' ���肵�������R�[�h��String(UNICODE)�ɕϊ�
    Dim uni As String
    Select Case cod
        Case "SJIS"
            ' SJIS����UNICODE�ւ̕ϊ�
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case "JIS"
            ' JIS����UNICODE�ւ̕ϊ�
            uni = objNonCode.JIS_To_VbUnicode(idat)
        Case "EUC"
            ' EUC����UNICODE�ւ̕ϊ�
            uni = objNonCode.EUC_To_VbUnicode(idat)
        Case "UNICODE"
            ' UNICODE����UNICODE�ւ̕ϊ�
            uni = objNonCode.UNICODE_To_VbUnicode(idat)
        Case "UTF7"
            ' UTF-7����UNICODE�ւ̕ϊ�
            uni = objNonCode.UTF7_To_VbUnicode(idat)
        Case "UTF8"
            ' UTF-8����UNICODE�ւ̕ϊ�
            uni = objNonCode.UTF8_To_VbUnicode(idat)
        Case "BIN"
            ' SJIS����UNICODE�ւ̕ϊ�
            uni = objNonCode.SJIS_To_VbUnicode(idat)
        Case Else
            ' SJIS����UNICODE�ւ̕ϊ�
            uni = objNonCode.SJIS_To_VbUnicode(idat)
    End Select

    ' �Ǎ��t�@�C���̉��s�R�[�h��CRLF�֕ϊ�
    uni = objNonCode.ChangeReturnToCrLf(uni)
    
    ' String(UNICODE)���o�͂����������R�[�h��Byte�z��ɕϊ�
    Dim odat() As Byte
    cod = strOutCode
    Select Case cod
        Case "SJIS"
            ' UNICODE����SJIS�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_SJIS(uni)
        Case "JIS"
            ' UNICODE����JIS�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_JIS(uni)
        Case "EUC"
            ' UNICODE����EUC�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_EUC(uni)
        Case "UNICODE"
            ' UNICODE����UNICODE�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_UNICODE(uni)
        Case "UTF7"
            ' UNICODE����UTF7�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_UTF7(uni)
        Case "UTF8"
            ' UNICODE����UTF8�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_UTF8(uni)
        Case Else
            ' UNICODE����SJIS�ւ̕ϊ�
            odat = objNonCode.VbUnicode_To_SJIS(uni)
    End Select

    ' �o�̓t�@�C�����o�C�i���`���ŏo��
    Dim opath As String: opath = pSrc & ".txt" 'App.Path & "\TestOut.txt"
    If Len(Dir(opath)) <> 0 Then
        Kill opath
    End If
    intFileNo = FreeFile
    Open opath For Binary As intFileNo
    Put intFileNo, , odat
    Close intFileNo
End Sub

Private Sub optMode_Click(Index As Integer)
    cmdCmpi.Enabled = True
End Sub


Private Sub nankan2Txt(src As String, file As String, wfile As String)

Dim objRegExp As RegExp 'RegExp�F[�Q�Ɛݒ�]�� Microsoft VBScript Regular Expressions 5.5 �Ƀ`�F�b�N��t����
Dim strResult As String '�u����̕�����
Dim Matches
Dim Match
Dim fn As Integer
Dim wfn As Integer
Dim lCnt As Integer
Dim data() As String
Dim wk As String
Dim wk2 As String
Dim wkPrt As String
Dim pos As Long
Dim phase As Long
Dim raceNo As String
Dim wkRaceNo As String
Dim retstr As String '
Dim nen As String
Dim gatu As String
Dim niti As String
Dim basho As String
Dim cmpininki As String
Dim cmpidata(12, 20) As String      'ninki,value
Dim umaban As Integer
Dim value As String
Dim idx As Integer
Dim backup As String
Dim wakCnt As Integer
Dim wakD() As String
Dim plc As Integer
Dim kire As String
Dim smpl As String
Dim wban As String
Dim cmpV As String
Dim wkstr As String
Dim dptr As Integer
Dim cutstr As String
Dim tmp As String

cmdCmpi.Enabled = False

Dim wkwk As String
    If optMode(0).value = True Then
         wkwk = "�n��"
    Else
         wkwk = "wakuNum"
    End If


'HTML�t�@�C��(param.)���������[�ɓW�J
'<<�t�@�C�� �J>>
fn = FreeFile
Open src For Input As #fn

'<<�t�@�C�� ��>>
lCnt = 0
Do Until EOF(fn)
    Line Input #fn, wk
    ReDim Preserve data(lCnt)
    data(lCnt) = wk
    lCnt = lCnt + 1
Loop

'<<�t�@�C�� ��>>
Close #fn

wfn = FreeFile
Open wfile For Append As #wfn


'<<�f�[�^���>>
'���K�\���I�u�W�F�N�g�̐錾
Set objRegExp = New RegExp

With objRegExp
    .Global = True '�����}�b�`��
    .IgnoreCase = True
    .Global = True
    
    phase = 0
    For lCnt = 0 To UBound(data)
        
        Select Case phase
        '�J�Ïꏊ�A�N����������
        '<TH BGCOLOR="#F56403" COLSPAN=31><FONT SIZE=+2>�n�ԃR���s�@�@�@�@</FONT><FONT SIZE=+2>2008�N1��19�� 1�񒆎R5����</FONT><FONT SIZE=+2>�@�@�@�@�g�ԃR���s</FONT></TH>
        '<h2 id="contentTit">2012�N1��5���@�R���s�w���|1�񒆎R1����</h2>
        Case 0
'             .Pattern = "<FONT SIZE=\+2>20.+����"
             .Pattern = "<h2 id=""contentTit"">.+����"
             
            pos = 0
            Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
            For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
               pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
               retstr = Match.value
            Next
            If pos = 0 Then
'                .Pattern = "<font size=""\+2"">20.+����"
                .Pattern = "contentTit""\>20.+����"
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
            End If
            
            If pos <> 0 Then
                '<FONT SIZE=+2>2008�N1��20�� 1�񒆎R6����
                '�N
                 .Pattern = ">.+�N"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                nen = Mid$(wk, 2, 4)
                '��
                 .Pattern = "�N.+��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 3 Then
                    gatu = Mid$(wk, 2, 1)
                Else
                    gatu = Mid$(wk, 2, 2)
                End If
                '��
                 .Pattern = "��.+��.+�R���s�w��"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                If Len(wk) = 16 Then
                    niti = Mid$(wk, 2, 1)
                Else
                    niti = Mid$(wk, 2, 2)
                End If
                '�J�Ïꏊ
                 .Pattern = "��.+����"
                pos = 0
                Set Matches = .Execute(retstr)   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   wk = Match.value
                Next
                Select Case Mid$(wk, 2, 2)
                Case "���"
                    basho = "45"
                Case "�D��"
                    basho = "43"
                Case "�Y�a"
                    basho = "42"
                Case "���"
                    basho = "44"
                End Select
                
                phase = 1
            End If
        Case 1
            Do
                If InStr(data(lCnt), "�y�g�ԃR���s�z") > 0 Then
                    Close #wfn
                    
                    cmdCmpi.Enabled = True
                    
                    Exit Sub
                End If
                
                .Pattern = "race""\>.+R\<\/span\>"
                 
                pos = 0
                Set Matches = .Execute(data(lCnt))   ' ���������s���܂��B
                For Each Match In Matches   ' Matches �R���N�V�����ɑ΂��ČJ��Ԃ��������s���܂��B
                   pos = Match.FirstIndex       '��v���镶���񂪌��������ʒu
                   retstr = Match.value
                Next
                
                If pos <> 0 Then
                    If Mid$(retstr, 8, 1) = "R" Then
                        tmp = Mid$(retstr, 7, 1)
                    Else
                        tmp = Mid$(retstr, 7, 2)
                    End If
                    raceNo = tmp
                    
                    wkstr = nen & Format$(gatu, "00") & Format$(niti, "00") & basho & Format$(raceNo, "00")
                    
                    '�f�[�^�擾
                    cutstr = data(lCnt)
                    Do
                        dptr = InStr(cutstr, "<br />")
                        If dptr = 0 Then
                            Print #wfn, wkstr
                            Exit Do
                        End If
                        
                        If Mid$(cutstr, dptr - 2, 1) = ">" Then
                            tmp = Mid$(cutstr, dptr - 1, 1)
                        Else
                            tmp = Mid$(cutstr, dptr - 2, 2)
                        End If
                        wkstr = wkstr & "," & tmp
                        If Mid$(cutstr, dptr + 7, 1) = "<" Then
                            tmp = Mid$(cutstr, dptr + 6, 1)
                        Else
                            tmp = Mid$(cutstr, dptr + 6, 2)
                        End If
                        
                        wkstr = wkstr & "," & tmp
                        
                        cutstr = Mid$(cutstr, dptr + 10)
                    Loop
                    
                    
                End If
                lCnt = lCnt + 1
            Loop
        End Select
        
    Next lCnt
    
End With

Set objRegExp = Nothing

'�e�L�X�g�t�@�C��(param.)�֏o��
src = file
fn = FreeFile
Open src For Append As #fn

'<<�t�@�C�� ��>>

wk = Format$(nen, "0000") & Format$(gatu, "00") & Format$(niti, "00") & basho

For idx = 1 To 12
    wk2 = ""
    For lCnt = 1 To 20
        wk2 = wk2 & "," & cmpidata(idx, lCnt)
    Next lCnt
    
    wk2 = wk & Format$(idx, "00") & wk2
    Print #fn, wk2

Next idx

'<<�t�@�C�� ��>>
Close #fn
Close #wfn

cmdCmpi.Enabled = True

End Sub

