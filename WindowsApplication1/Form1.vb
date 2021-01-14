
Imports system.Runtime.InteropServices

Public Class Form1
    Private Const mSnapOffset As Integer = 35
    Private Const WM_WINDOWPOSCHANGING As Integer = &H46

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure WINDOWPOS
        Public hwnd As IntPtr
        Public hwndInsertAfter As IntPtr
        Public x As Integer
        Public y As Integer
        Public cx As Integer
        Public cy As Integer
        Public flags As Integer
    End Structure

    Protected Overrides Sub WndProc(ByRef m As Message)
        ' Listen for operating system messages
        Select Case m.Msg
            Case WM_WINDOWPOSCHANGING
                SnapToDesktopBorder(Me, m.LParam, 0)
        End Select

        MyBase.WndProc(m)
    End Sub

    Public Shared Sub SnapToDesktopBorder(ByVal clientForm As Form, ByVal LParam As IntPtr, ByVal widthAdjustment As Integer)
        If clientForm Is Nothing Then
            ' Satisfies rule: Validate parameters
            Throw New ArgumentNullException("clientForm")
        End If

        ' Snap client to the top, left, bottom or right desktop border
        ' as the form is moved near that border.

        Try
            ' Marshal the LPARAM value which is a WINDOWPOS struct
            Dim NewPosition As New WINDOWPOS
            NewPosition = CType(Runtime.InteropServices.Marshal.PtrToStructure( _
                LParam, GetType(WINDOWPOS)), WINDOWPOS)

            If NewPosition.y = 0 OrElse NewPosition.x = 0 Then
                Return ' Nothing to do!
            End If

            ' Adjust the client size for borders and caption bar
            Dim ClientRect As Rectangle = clientForm.RectangleToScreen(clientForm.ClientRectangle)
            ClientRect.Width += SystemInformation.FrameBorderSize.Width - widthAdjustment
            ClientRect.Height += (SystemInformation.FrameBorderSize.Height + SystemInformation.CaptionHeight)

            ' Now get the screen working area (without taskbar)
            Dim WorkingRect As Rectangle = Screen.GetWorkingArea(clientForm.ClientRectangle)

            ' Left border
            If NewPosition.x >= WorkingRect.X - mSnapOffset AndAlso _
                NewPosition.x <= WorkingRect.X + mSnapOffset Then
                NewPosition.x = WorkingRect.X
            End If

            ' Get screen bounds and taskbar height (when taskbar is horizontal)
            Dim ScreenRect As Rectangle = Screen.GetBounds(Screen.PrimaryScreen.Bounds)
            Dim TaskbarHeight As Integer = ScreenRect.Height - WorkingRect.Height

            ' Top border (check if taskbar is on top or bottom via WorkingRect.Y)
            If NewPosition.y >= -mSnapOffset AndAlso _
                (WorkingRect.Y > 0 AndAlso NewPosition.y <= (TaskbarHeight + mSnapOffset)) OrElse _
                (WorkingRect.Y <= 0 AndAlso NewPosition.y <= (mSnapOffset)) Then
                If TaskbarHeight > 0 Then
                    NewPosition.y = WorkingRect.Y ' Horizontal Taskbar
                Else
                    NewPosition.y = 0 ' Vertical Taskbar
                End If
            End If

            ' Right border
            If NewPosition.x + ClientRect.Width <= WorkingRect.Right + mSnapOffset AndAlso _
                NewPosition.x + ClientRect.Width >= WorkingRect.Right - mSnapOffset Then
                NewPosition.x = WorkingRect.Right - (ClientRect.Width + SystemInformation.FrameBorderSize.Width)
            End If

            ' Bottom border
            If NewPosition.y + ClientRect.Height <= WorkingRect.Bottom + mSnapOffset AndAlso _
                NewPosition.y + ClientRect.Height >= WorkingRect.Bottom - mSnapOffset Then
                NewPosition.y = WorkingRect.Bottom - (ClientRect.Height + SystemInformation.FrameBorderSize.Height)
            End If

            ' Marshal it back
            Runtime.InteropServices.Marshal.StructureToPtr(NewPosition, LParam, True)
        Catch ex As ArgumentException
        End Try
    End Sub
End Class
