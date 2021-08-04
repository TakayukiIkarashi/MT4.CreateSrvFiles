Option Explicit

'MT4のサーバーリスト
Const MT4_SERVER_LIST = "https://www.trade-copier.com/index.php/features/supported-brokers/mt4-brokers-list"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim sh
Set sh = CreateObject("WScript.Shell")

'メイン処理実行
Call Main()

'メイン処理
Sub Main()
    '警告メッセージ
    Dim msg
    msg = ""
    msg = msg & "<<注意！>>" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "事前にMetaTraderを起動し、「ファイル(F)」メニューより「デモ口座の申請(A)」を選択し、"
    msg = msg & "「デモ口座の申請」ウィンドウを開いておいてください。" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "また、MetaTrader以外のすべてのアプリケーションを閉じてください。" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "このスクリプトを実行後、10秒以内に「デモ口座の申請」ウィンドウを最前面にしておいてください。"
    Call MsgBox(msg, vbExclamation + vbOkOnly)

    Dim isCrtLst
    isCrtLst = True
    msg = ""
    msg = msg & "ブローカーリストを更新しますか？"
    If (MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes) Then
        isCrtLst = False
    End If

    msg = ""
    msg = msg & "実行します。よろしいですか？"
    If (MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes) Then
        Exit Sub
    End If

    'スクリプト強制停止ファイルを起動
    sh.Run GetParentDir() & "StopWScript.exe"

    '10秒以内にMetaTraderを起動し、「デモ口座の申請」ウィンドウを開いておいてもらう
    WScript.Sleep 10000

    'すべてのIEのプロセスを強制終了
    Call TerminateIE()

    Dim brklst
    brklst = GetListfilePath()

    'ブローカーリスト生成
    If (isCrtLst) Then
        If (CreateBrokerList(brklst) = False) Then
            Exit Sub
        End If
    End If

    'ブローカーリスト読み込み
    Dim brokers
    Call ReadBrokerList(brklst, brokers)

    'ブローカーリストの内容を、1件ずつ「デモ口座の申請」ウィンドウに打ち込み
    Dim i
    For i = 0 To UBound(brokers)
        Dim brk
        brk = brokers(i)

        sh.SendKeys("{Enter}")
        WScript.Sleep 500

        Dim j
        For j = 1 To Len(brk)
            Dim c
            c = Mid(brk, j, 1)

            sh.SendKeys("{" & c & "}")

            WScript.Sleep 10
        Next

        sh.SendKeys("{Enter}")

        WScript.Sleep 3000
    Next
End Sub

'IEをすべて強制終了する
Sub TerminateIE()
    Const PROC_NAME = "iexplore.exe"

    On Error Resume Next

    Dim procList
    Set procList = GetObject("winmgmts:").InstancesOf("win32_process")

    Dim p
    For Each p In procList
        If (LCase(p.Name) = PROC_NAME) Then
            p.Terminate
        End If
    Next

    If (Err.Number <> 0) Then
        Err.Clear
    End If

    On Error GoTo 0
End Sub

'ブローカーリストを作成
Function CreateBrokerList(ByVal brklst)
    CreateBrokerList = False

    'ネットからブローカーリストを取得
    Dim s
    s = GetHTML()
    If (s = "") Then
        Call MsgBox("ブローカーリストの取得に失敗しました。", vbCritical + vbOkOnly)
        Exit Function
    End If

    '開始位置
    Dim ps
    ps = InStr(1, s, "<br>")
    Do
        ps = ps - 1

        If (ps < 1) Then
            Call MsgBox("ブローカーリストの取得先のHTMLに変更があったようです。", vbCritical + vbOkOnly)
            Exit Function
        End If

        If (Mid(s, ps, 1) = ">") Then
            ps = ps + 1
            Exit Do
        End If
    Loop

    '終了位置
    Dim pe
    pe = InStr(ps, s, "</span>")

    Dim r
    r = Mid(s, ps, pe - ps)
    r = Replace(r, "<br>", vbCrLf)

    'ブローカーリストを書き込み
    Dim fw
    Set fw = fso.OpenTextFile(brklst, 2)
    Call fw.Write(r)
    Call fw.Close()

    CreateBrokerList = True
End Function

'ブローカーリストのHTMLを取得
Function GetHTML()
    On Error Resume Next

    Dim ie
    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = False

    ie.Navigate(MT4_SERVER_LIST)

    Do
        If (ie.Busy = False) Then
            Exit Do
        End If

        WScript.Sleep 100
    Loop

    Dim html
    html = ie.Document.All(0).InnerHtml

    Call ie.Quit()

    If (Err.Number <> 0) Then
        MsgBox CStr(Err.Number) & ": " & Err.Description
        GetHTML = ""
        Exit Function
    End If

    GetHTML = html
End Function

'ブローカーリストを読み込み、配列に格納
Sub ReadBrokerList(ByVal fp, ByRef brokers)
    Dim f
    Set f = fso.OpenTextFile(fp)

    Dim s
    s = f.ReadAll()

    f.Close()

    brokers = Split(s, vbCrLf)
End Sub

'ブローカーリストのパスを返す
Function GetListfilePath()
    GetListfilePath = GetParentDir() & "broker.lst"
End Function

'このスクリプトファイルが存在するディレクトリのパスを返す
Function GetParentDir()
    GetParentDir = fso.getParentFolderName(WScript.ScriptFullName) & "\"
End Function
