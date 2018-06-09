Attribute VB_Name = "Module1"
Public Const GC_APLI_NAME = "O.K.馬"
Public Const GC_THANKS = "をご利用いただき、ありがとうございます！"
Public Const GC_AMAZON = "<a href="""">Amazonでのお買い物は、こちらから！開発支援にご協力お願いします。</a>"
Public Const GC_BLOG_MAIL = "a585c4de0e448f@mo.jugem.jp"
Public Const GC_MAC_MAIL = "racesoft@buhi-buhi.com"
Public Const GC_FAIL_MAIL = "利用確認メールが送信できませんでした。"
Public ws As DAO.Workspace
Public db As DAO.Database
Public Rs As DAO.Recordset
Public Rs3 As DAO.Recordset
Public ws2 As DAO.Workspace
Public db2 As DAO.Database
Public Rs2 As DAO.Recordset
Public RsWk As DAO.Recordset

Public gDB As String
Public fn As Long

Public gstrSql As String

Public gYear                As String
Public gCheckRace As String
Public gTime                As Integer
Public gMonth As String
Public gDay As String
Public gMonthDay            As String
Public gGetDetailMonthDayFlag            As String
Public gJyoCD               As String
Public gSyussoTosu As String
Public gSyussoTosuArr() As String
Public gKaiji As String
Public gNichiji As String
Public gRaceNum             As String

Public gFilename(5) As String
