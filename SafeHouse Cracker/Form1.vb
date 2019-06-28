Imports System.Runtime.InteropServices
Public Class Form1
    Dim Str_List As New List(Of String)
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowText(hWnd As IntPtr, text As Text.StringBuilder, count As Integer) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetWindowTextLength(hWnd As IntPtr) As Integer
    End Function
    Sub About()
        MsgBox("Open the SafeHouse file to be cracked to begin!" & vbCrLf & "Keep that window focused because keys are constantly sent to it." & vbCrLf & vbCrLf & "Icons made by Freepik from www.flaticon.com are licensed by CC BY 3.0" & vbCrLf & vbCrLf & "This program is not endorsed or affiliated in any way to SafeHouse™ or to Freepik.")
    End Sub
    Function BruteForce(ByVal Position As Integer, ByVal List_String As List(Of String)) As String
        Dim Current_Position = Position
        Dim Mod_List As New List(Of Decimal)
        Dim Output As String = Nothing

        While Current_Position >= 0
            Mod_List.Add(Current_Position Mod List_String.Count)
            Current_Position -= Current_Position Mod List_String.Count
            Current_Position /= List_String.Count
            Current_Position -= 1
        End While
        For i = (Mod_List.Count - 1) To 0 Step -1
            Output += List_String(Mod_List(i))
        Next
        Return Output
    End Function

    Private Function GetCaptionOfActiveWindow() As String
        Dim strTitle As String = String.Empty
        Dim handle As IntPtr = GetForegroundWindow()
        Dim stringBuilder As New Text.StringBuilder(GetWindowTextLength(handle) + 1)
        If GetWindowText(handle, stringBuilder, GetWindowTextLength(handle) + 1) > 0 Then
            strTitle = stringBuilder.ToString()
        End If
        Return strTitle
    End Function
    Dim TotalC = -1
    Dim S
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        While GetCaptionOfActiveWindow() = "Open SafeHouse Volume"
            TotalC = TextBox1.Text + 1
            TextBox1.Text = TotalC
            S = BruteForce(TotalC, Str_List)
            Clipboard.SetText(S)
            SendKeys.SendWait("^v{Enter}{Enter}")
        End While
        Label2.Text = "Password Tried: " & S
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        About()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Str_List.Clear()
        For Each C As Char In "0123456789"
            Str_List.Add(C.ToString)
        Next
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Str_List.Clear()
        For Each C As Char In "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            Str_List.Add(C.ToString)
        Next
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Str_List.Clear()
        For Each C As Char In "abcdefghijklmnopqrstuvwxyz"
            Str_List.Add(C.ToString)
        Next
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Str_List.Clear()
        For Each C As Char In "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
            Str_List.Add(C.ToString)
        Next
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged, TextBox2.TextChanged
        Str_List.Clear()
        For Each C As Char In TextBox2.Text
            Str_List.Add(C.ToString)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = -1
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If e.KeyCode = Keys.F1 Then About()
    End Sub
End Class
