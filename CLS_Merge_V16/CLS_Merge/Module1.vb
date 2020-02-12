Option Strict Off
Option Explicit Off
Imports VB = Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Text

Module Module1

    '### Shellwait ###
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredaccess As Integer, ByVal bInherithandle As Integer, ByVal dwProcessid As Integer) As Integer
    Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Integer, ByRef lpexitcode As Integer) As Integer
    Const STILL_ACTIVE As Short = &H103S
    Const PROCESS_QUERY_INFORMATION As Short = &H400S

    Public InVoer As String
    Public UitVoer As String
    Public LoCatie As String
    Public EDITOR As String
    Public Help As String
    Public Hoogte As Integer
    Public Scherm As Integer

    Sub ShellWait(ByRef sCommandLine As String)

        Dim hShell As Integer
        Dim hProc As Integer
        Dim lExit As Integer
        hShell = Shell(sCommandLine, AppWinStyle.NormalFocus)
        hProc = OpenProcess(PROCESS_QUERY_INFORMATION, False, hShell)

        Do
            GetExitCodeProcess(hProc, lExit)

            System.Windows.Forms.Application.DoEvents() : Sleep(100)

        Loop While lExit = STILL_ACTIVE
    End Sub

    Sub StartOff()
        Dim cmdline As String
        Dim noinput As String
        Help = "K:\Systems\Apps\CLS_Merge\ReadMe.pdf"
        cmdline = VB.Command()

        If File.Exists("C:\Program Files\Notepad++\notepad++.exe") Then
            EDITOR = "C:\Program Files\Notepad++\notepad++.exe"
        ElseIf File.Exists("C:\Program Files (x86)\Notepad++\notepad++.exe") Then
            EDITOR = "C:\Program Files (x86)\Notepad++\notepad++.exe"
        Else
            EDITOR = "C:\Windows\notepad.exe"
        End If
        If cmdline = "" Then
            noinput = MsgBox("Programma moet met Koppel file als input gestart worden", MsgBoxStyle.Exclamation)
            System.Diagnostics.Process.Start(Help)
            End
        End If
        InVoer = Right(cmdline, Len(cmdline) - InStrRev(cmdline, "\"))
        UitVoer = Left(InVoer, Len(InVoer) - (Len(InVoer) - InStrRev(InVoer, ".dat"))) & "cls"
        LoCatie = Left(cmdline, Len(cmdline) - (Len(cmdline) - InStrRev(cmdline, "\")))

        Form1.Label1.Text = InVoer
        Form1.Label2.Text = UitVoer
    End Sub
End Module
