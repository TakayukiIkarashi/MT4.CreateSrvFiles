Option Explicit

'MT4�̃T�[�o�[���X�g
Const MT4_SERVER_LIST = "https://www.trade-copier.com/index.php/features/supported-brokers/mt4-brokers-list"

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim sh
Set sh = CreateObject("WScript.Shell")

'���C���������s
Call Main()

'���C������
Sub Main()
    '�x�����b�Z�[�W
    Dim msg
    msg = ""
    msg = msg & "<<���ӁI>>" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "���O��MetaTrader���N�����A�u�t�@�C��(F)�v���j���[���u�f�������̐\��(A)�v��I�����A"
    msg = msg & "�u�f�������̐\���v�E�B���h�E���J���Ă����Ă��������B" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "�܂��AMetaTrader�ȊO�̂��ׂẴA�v���P�[�V��������Ă��������B" & vbCrLf
    msg = msg & vbCrLf
    msg = msg & "���̃X�N���v�g�����s��A10�b�ȓ��Ɂu�f�������̐\���v�E�B���h�E���őO�ʂɂ��Ă����Ă��������B"
    Call MsgBox(msg, vbExclamation + vbOkOnly)

    Dim isCrtLst
    isCrtLst = True
    msg = ""
    msg = msg & "�u���[�J�[���X�g���X�V���܂����H"
    If (MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes) Then
        isCrtLst = False
    End If

    msg = ""
    msg = msg & "���s���܂��B��낵���ł����H"
    If (MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes) Then
        Exit Sub
    End If

    '�X�N���v�g������~�t�@�C�����N��
    sh.Run GetParentDir() & "StopWScript.exe"

    '10�b�ȓ���MetaTrader���N�����A�u�f�������̐\���v�E�B���h�E���J���Ă����Ă��炤
    WScript.Sleep 10000

    '���ׂĂ�IE�̃v���Z�X�������I��
    Call TerminateIE()

    Dim brklst
    brklst = GetListfilePath()

    '�u���[�J�[���X�g����
    If (isCrtLst) Then
        If (CreateBrokerList(brklst) = False) Then
            Exit Sub
        End If
    End If

    '�u���[�J�[���X�g�ǂݍ���
    Dim brokers
    Call ReadBrokerList(brklst, brokers)

    '�u���[�J�[���X�g�̓��e���A1�����u�f�������̐\���v�E�B���h�E�ɑł�����
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

'IE�����ׂċ����I������
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

'�u���[�J�[���X�g���쐬
Function CreateBrokerList(ByVal brklst)
    CreateBrokerList = False

    '�l�b�g����u���[�J�[���X�g���擾
    Dim s
    s = GetHTML()
    If (s = "") Then
        Call MsgBox("�u���[�J�[���X�g�̎擾�Ɏ��s���܂����B", vbCritical + vbOkOnly)
        Exit Function
    End If

    '�J�n�ʒu
    Dim ps
    ps = InStr(1, s, "<br>")
    Do
        ps = ps - 1

        If (ps < 1) Then
            Call MsgBox("�u���[�J�[���X�g�̎擾���HTML�ɕύX���������悤�ł��B", vbCritical + vbOkOnly)
            Exit Function
        End If

        If (Mid(s, ps, 1) = ">") Then
            ps = ps + 1
            Exit Do
        End If
    Loop

    '�I���ʒu
    Dim pe
    pe = InStr(ps, s, "</span>")

    Dim r
    r = Mid(s, ps, pe - ps)
    r = Replace(r, "<br>", vbCrLf)

    '�u���[�J�[���X�g����������
    Dim fw
    Set fw = fso.OpenTextFile(brklst, 2)
    Call fw.Write(r)
    Call fw.Close()

    CreateBrokerList = True
End Function

'�u���[�J�[���X�g��HTML���擾
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

'�u���[�J�[���X�g��ǂݍ��݁A�z��Ɋi�[
Sub ReadBrokerList(ByVal fp, ByRef brokers)
    Dim f
    Set f = fso.OpenTextFile(fp)

    Dim s
    s = f.ReadAll()

    f.Close()

    brokers = Split(s, vbCrLf)
End Sub

'�u���[�J�[���X�g�̃p�X��Ԃ�
Function GetListfilePath()
    GetListfilePath = GetParentDir() & "broker.lst"
End Function

'���̃X�N���v�g�t�@�C�������݂���f�B���N�g���̃p�X��Ԃ�
Function GetParentDir()
    GetParentDir = fso.getParentFolderName(WScript.ScriptFullName) & "\"
End Function
