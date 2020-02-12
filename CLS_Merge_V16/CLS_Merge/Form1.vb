Option Strict Off
Imports System
Imports System.IO
Imports System.Text
Public Class Form1
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StartOff()
        Label3.Text = ""
        Label7.Text = ""
        Label9.Text = ""

        Button1.Visible = False
        LogBox.Visible = False

        Me.Width = 765
        Me.Left = 0
        Me.Top = 0
        Me.Show()
        Me.Refresh()
        Koppel()

        'System.Windows.Forms.Application.DoEvents() : Sleep(10000)
        'End
    End Sub
    Private Sub Koppel()
        Dim s1 As String
        Dim s2 As String
        Dim s As String
        Dim aantal As Integer = 0
        Dim aantalsub As Integer
        Dim subjes As Integer = 0
        Dim fout As Boolean
        Dim warning As Boolean
        Dim sw As StreamWriter
        Dim RenumBer As Boolean = False
        Dim RenumBered As Boolean = False
        Dim T1(999) As Integer
        Dim T2(999) As Integer
        Dim van As Integer = 0
        Dim naar As Integer = 0
        Dim index As Integer = 0


        On Error GoTo foutdelete
        If File.Exists(LoCatie & UitVoer) Then
            File.Delete(LoCatie & UitVoer)
            Label7.Text = "Start Koppelen... (Bestaande output wordt verwijderd)"
        Else
            Label7.Text = "Start Koppelen ..."
        End If
        '## lees het aantal subs en pas de hoogte aan ##'
        On Error GoTo foutinput
        FileOpen(1, LoCatie & InVoer, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        On Error GoTo foutlezenkoppel
        Do While Not EOF(1)
            s1 = LineInput(1)
            s1 = s1.ToLower
            If InStr(s1, "file = ") Then
                subjes = subjes + 1
            End If
            If InStr(s1, "replacetool = ") Then
                RenumBer = True
            End If
         Loop
        FileClose(1)

        Hoogte = Me.Height + (subjes * 15)
        Scherm = Screen.PrimaryScreen.Bounds.Height - 30
        If Hoogte > Scherm Then
            Me.Height = Scherm
        Else
            Me.Height = Me.Height + (subjes * 15)
        End If

        If RenumBer = True Then
            Label9.Text = "Gespecificeerde gereedschapnummers worden vervangen"
            Me.Width = 1000
            LogBox.Visible = True
            RenumBer = False
        End If

        '###############################################'
        On Error GoTo foutinput
        FileOpen(1, LoCatie & InVoer, OpenMode.Input, OpenAccess.Read, OpenShare.Shared)
        On Error GoTo foutlezenkoppel
        Do While Not EOF(1)
            s1 = LineInput(1)
            s2 = s1.ToLower
            '
            'replacetool = commando
            If InStr(s2, "replacetool = ") Then
                If InStr(s2, "yes") Then
                    RenumBer = True
                ElseIf InStr(s2, "no") Then
                    RenumBer = False
                Else
                    q1 = Split(s1, "=")

                    T1(index) = Int(Replace(q1(1), " ", ""))
                    T2(index) = Int(Replace(q1(2), " ", ""))

                    'sw = File.AppendText(LoCatie & UitVoer)
                    'sw.WriteLine("PPRINT/'(*** CLS Koppelprogramma >> T-nr: " & T1(index) & " Vervangen door T-nr: " & T2(index) & " ***)'")
                    'sw.Close()

                    index = index + 1
                    ReDim Preserve T1(index)
                    ReDim Preserve T2(index)
                End If
            End If
            '
            'insert = commando
            If InStr(s2, "insert = ") Then
                q1 = Replace(s2, "insert = ", "")
                q2 = q1.ToString.ToUpper
                sw = File.AppendText(LoCatie & UitVoer)
                sw.WriteLine(q2)
                sw.Close()
                aantal = aantal + 1
            End If
            '
            'file = commando
            If InStr(s2, "file = ") And RenumBer = False Then
                q1 = Replace(s1, "file = ", "")
                ListBox1.Items.Add("Koppel input file:")
                ListBox2.Items.Add(q1)
                Me.Refresh()
                aantalsub = 0
                On Error GoTo foutmetopenen
                Dim readText() As String = File.ReadAllLines(LoCatie & q1, System.Text.Encoding.Default)
                sw = File.AppendText(LoCatie & UitVoer)
                For Each s In readText
                    sw.WriteLine(s)
                Next
                sw.Close()
                aantal = aantal + readText.Length
                aantalsub = aantalsub + readText.Length
                ListBox3.Items.Add(" ... Gekoppeld. " & aantalsub & " Regels.")
                Me.Refresh()
            End If
            If InStr(s2, "file = ") And RenumBer = True Then
                RenumBered = True
                q1 = Replace(s2, "file = ", "")
                LogBox.Items.Add("Gereedschapnummers vervangen")
                ListBox1.Items.Add("Koppel input file:")
                ListBox2.Items.Add(q1)
                ListBox1.Items.Add(" en vervang GereedschapNr's")
                ListBox2.Items.Add("")
                Me.Refresh()
                aantalsub = 0
                aantalts = 0
                On Error GoTo foutmetopenen
                Dim readText() As String = File.ReadAllLines(LoCatie & q1, System.Text.Encoding.Default)
                sw = File.AppendText(LoCatie & UitVoer)
                For Each s In readText
                    If InStr(s, "LOAD/TOOL,") Or InStr(s, "TURRET/FACE,") Then
                        ww = Split(s, ",")
                        www = CInt(ww(1))
                        r = 0
                        newT = ww(1)
                        For Each p In T1
                            If www = p Then
                                newT = CStr(T2(r))
                            End If
                            r = r + 1
                        Next
                        If newT = ww(1) Then GoTo SKIPTHISONE
                        ww(1) = newT
                        LogBox.Items.Add(" > " & www & " vervangen door " & newT)
                        LogBox.SelectedIndex = LogBox.Items.Count - 1
                        s = Join(ww, ",")
                        i = "PPRINT/';CLS Koppel >> T-nr: " & www & " Vervangen door T-nr: " & newT & "'"
                        aantal = aantal + 1
                        sw.WriteLine(i)
                        aantalts = aantalts + 1
SKIPTHISONE:
                    End If
                    sw.WriteLine(s)
                Next
                sw.Close()
                aantal = aantal + readText.Length
                aantalsub = aantalsub + readText.Length
                ListBox3.Items.Add(" ... Gekoppeld. " & aantalsub & " Regels.")
                ListBox3.Items.Add(" ...   T-Nr's vervangen. " & aantalts & " x")
                Me.Refresh()
            End If
        Loop
        Label3.Text = "Koppelen is klaar!      " & aantal & " Regels geschreven."

        If RenumBered = True Then
            LogBox.Items.Add("klaar...")
            LogBox.SelectedIndex = LogBox.Items.Count - 1
        End If

        GoTo klaar

foutdelete:
        fout = True
        Label3.Text = "ERROR !!! Kan bestaande Output: " & UitVoer & " NIET verwijderen !!!"
        GoTo klaar

foutinput:
        fout = True
        Label3.Text = "ERROR !!! Kan Koppel file: " & InVoer & " NIET openen !!!"
        GoTo klaar

foutlezenkoppel:
        fout = True
        Label3.Text = "ERROR !!! Kan file: " & InVoer & " NIET lezen !!!"
        GoTo klaar

foutlezenaltts:
        fout = True
        Label3.Text = "ERROR !!! Kan file: " & ALT_Ts & " NIET lezen !!!"
        GoTo klaar

foutmetopenen:
        dd = Err.Description
        fout = False
        warning = True
        ListBox3.Items.Add("Kan Sub file NIET openen !!!")
        Label3.Text = "ERROR !!! Kan Koppel Sub file: " & q1 & " NIET openen !!!"
        sw.Close()
        FileClose(3)
        GoTo klaar

klaar:
        FileClose(1)
        FileClose(2)
        My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Hand)

        Label1.ForeColor = Color.Blue
        Label2.ForeColor = Color.Blue
        Me.Refresh()
        Button1.Visible = True
        Button4.Visible = True
        If fout Then
            If File.Exists(LoCatie & UitVoer) Then
                My.Computer.FileSystem.DeleteFile(LoCatie & UitVoer)
            End If
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Hand)
        Else
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Asterisk)
            If warning Then
                Dim Ret As MsgBoxResult
                Ret = MsgBox("CLS File NIET compleet!" & Chr(13) & "Niet alle sub files zijn gekoppeld" & Chr(13) & Chr(13) & "Delete CLS file?", MsgBoxStyle.YesNo, "CLS NIET compleet!")
                If Ret = MsgBoxResult.Yes Then
                    If File.Exists(LoCatie & UitVoer) Then
                        My.Computer.FileSystem.DeleteFile(LoCatie & UitVoer)
                    End If
                End If
            Else
                Button2.Visible = True
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Label2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.Click
        Dim ret As String
        ret = Shell(EDITOR & " " & Chr(34) & LoCatie & UitVoer & Chr(34), AppWinStyle.NormalFocus)
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click
        Dim ret As String
        'ret = Shell("c:\windows\notepad.exe " & Chr(34) & LoCatie & InVoer & Chr(34), AppWinStyle.NormalFocus)
        ret = Shell(EDITOR & " " & Chr(34) & LoCatie & InVoer & Chr(34), AppWinStyle.NormalFocus)
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.Height < 195 Then Me.Height = 195
        Label3.Top = Me.Height - 70
        ListBox1.Height = Me.Height - 200
        ListBox2.Height = Me.Height - 200
        ListBox3.Height = Me.Height - 200
        LogBox.Height = Me.Height - 64
        BackBox.Height = LogBox.Height + LogBox.Top - BackBox.Top


    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ret As String
        ChDir(LoCatie)
        'MsgBox(Chr(34) & Environ("localappdata") & "\DutchAero\PostProcessor V22.exe" & Chr(34) & " " & UitVoer)
        ret = Shell(Chr(34) & Environ("localappdata") & "\DutchAero\PostProcessor V22.exe" & Chr(34) & " " & UitVoer, AppWinStyle.NormalFocus)
        'End
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        On Error Resume Next
        System.Diagnostics.Process.Start(Help)

    End Sub
End Class
