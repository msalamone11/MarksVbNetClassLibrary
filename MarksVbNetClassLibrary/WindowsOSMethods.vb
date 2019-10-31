Imports System.Windows.Forms

Namespace WindowsOSMethods

    Public Module WindowsOSMethods

        Public Sub BringWindowToFront(WindowTitle As String)

            Dim ProcessList() As Process = Process.GetProcesses()
            Dim DesiredProgramsWindowTitle As String = Nothing

            For Each process As Process In ProcessList
                If String.IsNullOrEmpty(process.MainWindowTitle) = False Then
                    If process.MainWindowTitle.Contains(WindowTitle) Then
                        DesiredProgramsWindowTitle = process.MainWindowTitle
                        Exit For
                    End If
                End If
            Next

            SelectWindow(DesiredProgramsWindowTitle, Nothing)

        End Sub

        Public Sub SelectWindow(ByVal strWindowCaption As String, ByVal strClassName As String)

            Dim hWnd As Integer
            hWnd = FindWindow(strClassName, strWindowCaption)

            If hWnd > 0 Then
                SetForegroundWindow(hWnd)

                If IsIconic(hWnd) Then  'Restore if minimized
                    ShowWindow(hWnd, SW_RESTORE)
                Else
                    ShowWindow(hWnd, SW_SHOW)
                End If

            End If

        End Sub

        Public Const SW_RESTORE As Integer = 9 'Used by Select Window
        Public Const SW_SHOW As Integer = 5 'Used by Select Window
        Public Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer 'Used by Select Window
        Public Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer 'Used by Select Window
        Public Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Integer) As Boolean 'Used by Select Window
        Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer 'Used by Select Window

        Public Function GetScreenText()

            Clipboard.Clear()

            SelectAllText()
            CopyText()

            Return Clipboard.GetText

        End Function

        Public Sub SelectAllText()
            SendKeys.SendWait("^a")
        End Sub

        Public Sub CopyText()
            SendKeys.SendWait("^c")
        End Sub

    End Module

End Namespace