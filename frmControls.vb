Public Class frmControls

#Region "Buttons"

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Hide()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        frmInsteonControl.ProgressBar1.Value = frmInsteonControl.iRefresh - 2
        Me.Hide()
    End Sub

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        If frmInsteonControl.PowerLincModemConnect() Then
            btnDisconnect.Enabled = True
            btnConnect.Enabled = False
            btnRefresh.Enabled = True
            Me.Visible = True
            Me.Refresh()
            frmInsteonControl.bStartup = True
            frmInsteonControl.tEvents.Start()
            frmInsteonControl.InitializeTable()
        Else
            btnDisconnect.Enabled = False
            btnConnect.Enabled = True
            btnRefresh.Enabled = False
            frmInsteonControl.plm = Nothing
        End If

    End Sub

    Private Sub btnDisconnect_Click(sender As Object, e As EventArgs) Handles btnDisconnect.Click
        frmInsteonControl.tEvents.Stop()
        frmInsteonControl.tNewRequests.Stop()
        frmInsteonControl.tRefreshTable.Stop()
        frmInsteonControl.tPantry.Stop()
        frmInsteonControl.tMasterCloset.Stop()
        frmInsteonControl.tKidsBathroom.Stop()
        frmInsteonControl.tMudRoom.Stop()
        frmInsteonControl.plm = Nothing
        lblCOMConnectionStatus.Text = "Not Connected"
        btnDisconnect.Enabled = False
        btnConnect.Enabled = True
    End Sub
#End Region

#Region "Lost Focus"
    Private Sub txtRefreshRate_LostFocus(sender As Object, e As EventArgs) Handles mtxtRefreshRate.LostFocus
        Try
            frmInsteonControl.iRefresh = CInt(mtxtRefreshRate.Text)
            frmInsteonControl.ProgressBar1.Maximum = frmInsteonControl.iRefresh
            frmInsteonControl.SaveParameters()
        Catch ex As Exception
            mtxtRefreshRate.Text = frmInsteonControl.iRefresh.ToString
        End Try
    End Sub


    Private Sub mtxtDelay_LostFocus(sender As Object, e As EventArgs) Handles mtxtDelay.LostFocus
        Try
            If CInt(mtxtDelay.Text) < 1000 And CInt(mtxtDelay.Text) > 0 Then
                frmInsteonControl.iShortDelay = CInt(mtxtDelay.Text)
                frmInsteonControl.iLongDelay = frmInsteonControl.iShortDelay * 10
                frmInsteonControl.SaveParameters()
            Else
                mtxtDelay.Text = frmInsteonControl.iShortDelay.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub

    Private Sub txtCheckDB_LostFocus(sender As Object, e As EventArgs) Handles mtxtCheckDB.LostFocus
        Try
            If CInt(mtxtCheckDB.Text) >= 1000 Then
                frmInsteonControl.tNewRequests.Interval = CInt(mtxtCheckDB.Text)
                frmInsteonControl.SaveParameters()
            Else
                mtxtCheckDB.Text = frmInsteonControl.tNewRequests.Interval.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub

#End Region

End Class