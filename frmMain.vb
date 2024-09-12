Public Class frmInsteonControl
    Public WithEvents plm As Plm
    Const iNumDevices As Integer = 35
    Const iRetries As Integer = 10

    Structure TrackingList
        Dim DeviceID As String
        Dim DeviceName As String
        Dim DeviceState As Integer
        Dim DeviceEnabled As Boolean
        Dim DeviceDimmable As Boolean
        Dim DeviceActive As Boolean
    End Structure

    Dim Insteon(iNumDevices) As TrackingList
    Dim ExpectedStatus(iNumDevices) As Label
    Dim ActualStatus(iNumDevices) As Label
    Dim DisabledStatus(iNumDevices) As Label
    Dim OnButton(iNumDevices) As Button
    Dim OffButton(iNumDevices) As Button
    Dim DeviceLabel(iNumDevices) As Label
    Dim bLeakCheckSleep As Boolean
    Public bStartup As Boolean
    Public iRefresh As Integer = 300 'Seconds
    Public iShortDelay As Integer = 55
    Public iLongDelay As Integer = iShortDelay * 10

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            Event_HistoryTableAdapter1.InsertQuery("9028")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub frmInsteonControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim port As String
        Dim ports As String() = SerialPort.GetPortNames()
        Dim badChars As Char() = New Char() {"c"}

        Try
            Event_HistoryTableAdapter1.InsertQuery("9027")
            For Each port In ports
                ' .NET Framework has a bug where COM ports are
                ' sometimes appended with a 'c' characeter!
                If port.IndexOfAny(badChars) <> -1 Then
                    frmControls.cbCOMPort.Items.Add(port.Substring(0, port.Length - 1))
                Else
                    frmControls.cbCOMPort.Items.Add(port)
                End If
            Next

            If frmControls.cbCOMPort.Items.Count = 0 Then
                frmControls.cbCOMPort.Text = ""
            Else
                frmControls.cbCOMPort.Text = "COM4"
            End If

            Me.Visible = True
            Me.Refresh()
            bStartup = True
            bLeakCheckSleep = True

            If InitializeApp() Then
                frmControls.btnConnect.Enabled = False
                frmControls.btnDisconnect.Enabled = True
                frmControls.btnRefresh.Enabled = True
                InitializeTable()
                lblLastRefresh.Text = Now.ToString
            Else
                Throw New OperationCanceledException("Error Connecting")
            End If
        Catch ex As Exception
            plm = Nothing
            frmControls.btnConnect.Enabled = True
            frmControls.btnDisconnect.Enabled = False
            frmControls.btnRefresh.Enabled = False
            Insteon_HistoryTableAdapter1.InsertQuery("9002")
        End Try

    End Sub

#Region "Initialize"
    Public Sub SaveParameters()
        Dim FILE_NAME As String = "WDInsteonControlParameter.xml"
        Dim objWriter As New System.IO.StreamWriter(FILE_NAME)
        Dim strParameter As String

        strParameter = "<Parameter>"
        strParameter = strParameter & "<COM_Port>" & frmControls.cbCOMPort.Text & "</COM_Port>"
        strParameter = strParameter & "<Refresh_Rate>" & frmControls.mtxtRefreshRate.Text & "</Refresh_Rate>"
        strParameter = strParameter & "<Check_DB>" & frmControls.mtxtCheckDB.Text & "</Check_DB>"
        strParameter = strParameter & "<Delay>" & frmControls.mtxtDelay.Text & "</Delay>"
        strParameter = strParameter & "<Handle_Events>"
        If frmControls.ckHandleEvents.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Handle_Events>"
        strParameter = strParameter & "<Sydney_Room_Dim>" & mtxtSydneyRoomDim.Text & "</Sydney_Room_Dim>"
        strParameter = strParameter & "<Ethan_Room_Dim>" & mtxtEthanRoomDim.Text & "</Ethan_Room_Dim>"
        strParameter = strParameter & "<Master_Bedroom_Dim>" & mtxtMasterBedroomDim.Text & "</Master_Bedroom_Dim>"
        strParameter = strParameter & "<Spare_Bedroom_Dim>" & mtxtMasterBedroomDim.Text & "</Spare_Bedroom_Dim>"
        strParameter = strParameter & "<Pantry><Timer_Active>"
        If chkPantryTimer.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Timer_Active><Timer_Value>" & mtxtPantryTimer.Text & "</Timer_Value></Pantry>"
        strParameter = strParameter & "<Kids_Bathroom><Timer_Active>"
        If chkKidsBathroomTimer.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Timer_Active><Timer_Value>" & mtxtKidsBathroomTimer.Text & "</Timer_Value></Kids_Bathroom>"
        strParameter = strParameter & "<Master_Closet><Timer_Active>"
        If chkMasterClosetTimer.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Timer_Active><Timer_Value>" & mtxtMasterClosetTimer.Text & "</Timer_Value></Master_Closet>"
        strParameter = strParameter & "<Mud_Room><Timer_Active>"
        If chkMudRoomTimer.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Timer_Active><Timer_Value>" & mtxtMudRoomTimer.Text & "</Timer_Value></Mud_Room>"
        strParameter = strParameter & "<Basement_Stairs><Timer_Active>"
        If chkBasementStairsTimer.Checked Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Timer_Active><Timer_Value>" & mtxtBasementStairsTimer.Text & "</Timer_Value></Basement_Stairs>"

        strParameter = strParameter & "<Device_Active>"
        strParameter = strParameter & "<Master_Bedroom>"
        If Insteon(0).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Master_Bedroom>"
        strParameter = strParameter & "<Ethans_Light>"
        If Insteon(1).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Ethans_Light>"
        strParameter = strParameter & "<Sydneys_Light>"
        If Insteon(2).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Sydneys_Light>"
        strParameter = strParameter & "<Master_Fan>"
        If Insteon(3).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Master_Fan>"
        strParameter = strParameter & "<Ethans_Fan>"
        If Insteon(4).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Ethans_Fan>"
        strParameter = strParameter & "<Sydneys_Fan>"
        If Insteon(5).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Sydneys_Fan>"
        strParameter = strParameter & "<Master_Bathroom>"
        If Insteon(6).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Master_Bathroom>"
        strParameter = strParameter & "<Spare_Bedroom>"
        If Insteon(7).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Spare_Bedroom>"
        strParameter = strParameter & "<Kids_Hallway>"
        If Insteon(8).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Kids_Hallway>"
        strParameter = strParameter & "<Master_Closet>"
        If Insteon(9).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Master_Closet>"
        strParameter = strParameter & "<Spare_Bedroom_Fan>"
        If Insteon(10).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Spare_Bedroom_Fan>"
        strParameter = strParameter & "<Kids_Bathroom>"
        If Insteon(11).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Kids_Bathroom>"
        strParameter = strParameter & "<Mud_Room>"
        If Insteon(12).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Mud_Room>"
        strParameter = strParameter & "<Spider_Lamp>"
        If Insteon(13).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Spider_Lamp>"
        strParameter = strParameter & "<Pantry>"
        If Insteon(14).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Pantry>"
        strParameter = strParameter & "<Sconce_Light>"
        If Insteon(15).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Sconce_Light>"
        strParameter = strParameter & "<Living_Room>"
        If Insteon(16).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Living_Room>"
        strParameter = strParameter & "<Shell_Lamp>"
        If Insteon(17).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Shell_Lamp>"
        strParameter = strParameter & "<FD_Coach>"
        If Insteon(18).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</FD_Coach>"
        strParameter = strParameter & "<GD_Coach>"
        If Insteon(19).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</GD_Coach>"
        strParameter = strParameter & "<Backyard_Spot>"
        If Insteon(20).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Backyard_Spot>"
        strParameter = strParameter & "<Kitchen_Light>"
        If Insteon(21).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Kitchen_Light>"
        strParameter = strParameter & "<PlayRoom_East>"
        If Insteon(22).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</PlayRoom_East>"
        strParameter = strParameter & "<Front_Hall>"
        If Insteon(23).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Front_Hall>"
        strParameter = strParameter & "<Kitchen_Nook>"
        If Insteon(24).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Kitchen_Nook>"
        strParameter = strParameter & "<PlayRoom_West>"
        If Insteon(25).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</PlayRoom_West>"
        strParameter = strParameter & "<DiningRoom>"
        If Insteon(26).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</DiningRoom>"
        strParameter = strParameter & "<GarageInside>"
        If Insteon(27).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</GarageInside>"
        strParameter = strParameter & "<GDCoach2>"
        If Insteon(28).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</GDCoach2>"
        strParameter = strParameter & "<Basement_Stairs>"
        If Insteon(29).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Basement_Stairs>"
        strParameter = strParameter & "<Basement_N>"
        If Insteon(30).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Basement_N>"
        strParameter = strParameter & "<Art_Room>"
        If Insteon(31).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Art_Room>"
        strParameter = strParameter & "<Basement_Mid>"
        If Insteon(32).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Basement_Mid>"
        strParameter = strParameter & "<Server_Room>"
        If Insteon(33).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Server_Room>"
        strParameter = strParameter & "<Basement_S>"
        If Insteon(34).DeviceActive Then
            strParameter = strParameter & "True"
        Else
            strParameter = strParameter & "False"
        End If
        strParameter = strParameter & "</Basement_S>"
        'strParameter = strParameter & "<Lamp>"
        'If Insteon().DeviceActive Then
        '    strParameter = strParameter & "True"
        'Else
        '    strParameter = strParameter & "False"
        'End If
        'strParameter = strParameter & "</Lamp>"
        strParameter = strParameter & "</Device_Active>"

        strParameter = strParameter & "</Parameter>"

        objWriter.Write(strParameter)
        objWriter.Close()
    End Sub

    Private Sub LoadParameters()
        Dim xmlData As DataSet = New DataSet()
        Dim i As Integer

        'Load Parameters
        Try
            xmlData.ReadXml("WDInsteonControlParameter.xml")
            frmControls.cbCOMPort.Text = xmlData.Tables(0).Rows(0).ItemArray(0)
            frmControls.mtxtRefreshRate.Text = xmlData.Tables(0).Rows(0).ItemArray(1)
            iRefresh = CInt(frmControls.mtxtRefreshRate.Text)
            frmControls.mtxtCheckDB.Text = xmlData.Tables(0).Rows(0).ItemArray(2)
            frmControls.mtxtDelay.Text = xmlData.Tables(0).Rows(0).ItemArray(3)
            iShortDelay = CInt(frmControls.mtxtDelay.Text)
            tNewRequests.Interval = CInt(frmControls.mtxtCheckDB.Text)
            If xmlData.Tables(0).Rows(0).ItemArray(4) = "True" Then
                frmControls.ckHandleEvents.Checked = True
            Else
                frmControls.ckHandleEvents.Checked = False
            End If
            mtxtSydneyRoomDim.Text = xmlData.Tables(0).Rows(0).ItemArray(5)
            mtxtEthanRoomDim.Text = xmlData.Tables(0).Rows(0).ItemArray(6)
            mtxtMasterBedroomDim.Text = xmlData.Tables(0).Rows(0).ItemArray(7)
            mtxtSpareBedroomDim.Text = xmlData.Tables(0).Rows(0).ItemArray(8)
            If xmlData.Tables(1).Rows(0).Item(0) = "True" Then
                chkPantryTimer.Checked = True
            Else
                chkPantryTimer.Checked = False
            End If
            mtxtPantryTimer.Text = xmlData.Tables(1).Rows(0).Item(1)
            tPantry.Interval = CInt(mtxtPantryTimer.Text) * 1000 * 60

            If xmlData.Tables(2).Rows(0).Item(0) = "True" Then
                chkKidsBathroomTimer.Checked = True
            Else
                chkKidsBathroomTimer.Checked = False
            End If
            mtxtKidsBathroomTimer.Text = xmlData.Tables(2).Rows(0).Item(1)
            tKidsBathroom.Interval = CInt(mtxtKidsBathroomTimer.Text) * 1000 * 60

            If xmlData.Tables(3).Rows(0).Item(0) = "True" Then
                chkMasterClosetTimer.Checked = True
            Else
                chkMasterClosetTimer.Checked = False
            End If
            mtxtMasterClosetTimer.Text = xmlData.Tables(3).Rows(0).Item(1)
            tMasterCloset.Interval = CInt(mtxtMasterClosetTimer.Text) * 1000 * 60
            If xmlData.Tables(4).Rows(0).Item(0) = "True" Then
                chkMudRoomTimer.Checked = True
            Else
                chkMudRoomTimer.Checked = False
            End If
            mtxtMudRoomTimer.Text = xmlData.Tables(4).Rows(0).Item(1)
            tMudRoom.Interval = CInt(mtxtMudRoomTimer.Text) * 1000 * 60
            If xmlData.Tables(5).Rows(0).Item(0) = "True" Then
                chkBasementStairsTimer.Checked = True
            Else
                chkBasementStairsTimer.Checked = False
            End If
            mtxtBasementStairsTimer.Text = xmlData.Tables(5).Rows(0).Item(1)
            tBasementStairs.Interval = CInt(mtxtBasementStairsTimer.Text) * 1000 * 60
            For i = 0 To iNumDevices - 1
                If xmlData.Tables(6).Rows(0).Item(i) = "True" Then
                    Insteon(i).DeviceActive = True
                    DeviceLabel(i).ForeColor = SystemColors.ControlText
                Else
                    Insteon(i).DeviceActive = False
                    DeviceLabel(i).ForeColor = Color.Red
                End If
            Next

        Catch ex As Exception
            Insteon_HistoryTableAdapter1.InsertQuery("9002")
            frmControls.mtxtRefreshRate.Text = iRefresh.ToString
            frmControls.mtxtDelay.Text = iShortDelay.ToString
            tKidsBathroom.Interval = 30 * 1000 * 60
            mtxtKidsBathroomTimer.Text = "30"
            tPantry.Interval = 3 * 1000 * 60
            mtxtPantryTimer.Text = "3"
            tMasterCloset.Interval = 3 * 1000 * 60
            mtxtMasterClosetTimer.Text = "3"
            tMudRoom.Interval = 5 * 1000 * 60
            mtxtMudRoomTimer.Text = "5"
            frmControls.mtxtCheckDB.Text = tNewRequests.Interval.ToString
            mtxtBasementStairsTimer.Text = "3"
            tBasementStairs.Interval = 3 * 1000 * 60
            For i = 0 To iNumDevices - 1
                Insteon(i).DeviceActive = True
                DeviceLabel(i).ForeColor = SystemColors.ControlText
            Next
        End Try

        iLongDelay = iShortDelay * 10

    End Sub

    Public Function PowerLincModemConnect() As Boolean
        Dim oDevice As Object
        Dim i As Integer = 0
        Dim bInitialConnectSuccess As Boolean = False

        Try
            plm = Nothing
            plm = New Plm(frmControls.cbCOMPort.Text)

            'Try to confirm connection
            oDevice = Nothing
            While Not bInitialConnectSuccess And i <= iRetries
                bInitialConnectSuccess = plm.Network.TryConnectToDevice("1A.F4.47", oDevice)  'Spider Lamp
                i += 1
                Thread.Sleep(iLongDelay)
            End While
            oDevice = Nothing
            If Not bInitialConnectSuccess Then
                Throw New OperationCanceledException("Error Connecting")
            Else
                frmControls.lblCOMConnectionStatus.Text = "Connected"
            End If
            Return True
        Catch ex As Exception
            MsgBox("Error Connecting: " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9003")
            frmControls.lblCOMConnectionStatus.Text = "Not Connected"
            Return False
        End Try

    End Function

    Private Function InitializeApp() As Boolean
        'Runs once at startup

        Try
            DeviceLabel(0) = lblMasterBedroom
            DeviceLabel(1) = lblEthanBedroom
            DeviceLabel(2) = lblSydneyBedroom
            DeviceLabel(3) = lblMasterFan
            DeviceLabel(4) = lblEthanFan
            DeviceLabel(5) = lblSydneyFan
            DeviceLabel(6) = lblMasterBathroom
            DeviceLabel(7) = lblSpareBedroom
            DeviceLabel(8) = lblKidsHallway
            DeviceLabel(9) = lblMasterCloset
            DeviceLabel(10) = lblSpareBedroomFan
            DeviceLabel(11) = lblKidsBathroom
            DeviceLabel(12) = lblMudRoom
            DeviceLabel(13) = lblSpiderLamp
            DeviceLabel(14) = lblPantry
            DeviceLabel(15) = lblSconceLights
            DeviceLabel(16) = lblLivingRoom
            DeviceLabel(17) = lblSmartBulb
            DeviceLabel(18) = lblFDCoach
            DeviceLabel(19) = lblGDCoach1
            DeviceLabel(20) = lblBackyardSpot
            DeviceLabel(21) = lblKitchenLight
            DeviceLabel(22) = lblPlayRoom1
            DeviceLabel(23) = lblFrontHall
            DeviceLabel(24) = lblKitchenNook
            DeviceLabel(25) = lblPlayRoom2
            DeviceLabel(26) = lblDiningRoom
            DeviceLabel(27) = lblGarageInside
            DeviceLabel(28) = lblGDCoach2
            DeviceLabel(29) = lblBasementStairs
            DeviceLabel(30) = lblBasementN
            DeviceLabel(31) = lblArtRoom
            DeviceLabel(32) = lblBasementMid
            DeviceLabel(33) = lblServerRoom
            DeviceLabel(34) = lblBasementS
            'DeviceLabel() = lblLamp

            Insteon(0).DeviceID = "20.E5.84"  'Master Bedroom Light
            Insteon(0).DeviceName = "Master Bedroom Light"
            Insteon(0).DeviceDimmable = True
            Insteon(1).DeviceID = "23.B6.EB"  'Ethan Bedroom Light
            Insteon(1).DeviceName = "Ethan's Bedroom Light"
            Insteon(1).DeviceDimmable = True
            Insteon(2).DeviceID = "23.C5.30"  'Sydney's Bedroom Light
            Insteon(2).DeviceName = "Sydney's Bedroom Light"
            Insteon(2).DeviceDimmable = True
            Insteon(3).DeviceID = "2A.39.F6"  'Master Bedroom Fan
            Insteon(3).DeviceName = "Master Bedroom Fan"
            Insteon(3).DeviceDimmable = False
            Insteon(4).DeviceID = "2A.3B.F0"  'Ethan Bedroom Fan
            Insteon(4).DeviceName = "Ethan's Bedroom Fan"
            Insteon(4).DeviceDimmable = False
            Insteon(5).DeviceID = "2A.3D.80"  'Sydney Bedroom Fan
            Insteon(5).DeviceName = "Sydney's Bedroom Fan"
            Insteon(5).DeviceDimmable = False
            Insteon(6).DeviceID = "23.E2.E4"  'Master Bathroom
            Insteon(6).DeviceName = "Master Bathroom"
            Insteon(6).DeviceDimmable = False
            Insteon(7).DeviceID = "22.56.9E"  'Spare Bedroom Light
            Insteon(7).DeviceName = "Spare Bedroom Light"
            Insteon(7).DeviceDimmable = True
            Insteon(8).DeviceID = "29.2A.B9"  'Kids Hallway
            Insteon(8).DeviceName = "Kids Hallway"
            Insteon(8).DeviceDimmable = False
            Insteon(9).DeviceID = "23.EA.1D"  'Master Closet
            Insteon(9).DeviceName = "Master Closet"
            Insteon(9).DeviceDimmable = False
            Insteon(10).DeviceID = "2A.3F.41"  'Spare Bedroom Fan
            Insteon(10).DeviceName = "Spare Bedroom Fan"
            Insteon(10).DeviceDimmable = False
            Insteon(11).DeviceID = "23.DB.10"  'Kids Bathroom
            Insteon(11).DeviceName = "Kid's Bathroom"
            Insteon(11).DeviceDimmable = False


            Insteon(12).DeviceID = "29.27.CE"  'MudRoom
            Insteon(12).DeviceName = "Mud Room"
            Insteon(12).DeviceDimmable = False
            Insteon(13).DeviceID = "1A.F4.47"  'Spider Lamp
            Insteon(13).DeviceName = "Spider Lamp"
            Insteon(13).DeviceDimmable = False
            Insteon(14).DeviceID = "22.13.0C"  'Pantry
            Insteon(14).DeviceName = "Pantry"
            Insteon(14).DeviceDimmable = False
            Insteon(15).DeviceID = "29.27.9D"  'Sconce Lights
            Insteon(15).DeviceName = "Sconce Lights"
            Insteon(15).DeviceDimmable = False
            Insteon(16).DeviceID = "22.13.0E"  'Living Room
            Insteon(16).DeviceName = "Living Room"
            Insteon(16).DeviceDimmable = False
            Insteon(17).DeviceID = "20.25.3A"  'Shell Lamp
            Insteon(17).DeviceName = "Shell Lamp"
            Insteon(17).DeviceDimmable = False
            Insteon(18).DeviceID = "33.67.22"  'FD Coach Lights
            Insteon(18).DeviceName = "FD Coach Lights"
            Insteon(18).DeviceDimmable = False
            Insteon(19).DeviceID = "17.A6.A9"  'GD Coach Lights
            Insteon(19).DeviceName = "GD Coach Lights 1"
            Insteon(19).DeviceDimmable = False
            Insteon(20).DeviceID = "2B.82.42"  'Backyard Spot
            Insteon(20).DeviceName = "Back Yard Spot"
            Insteon(20).DeviceDimmable = False
            Insteon(21).DeviceID = "22.10.09"  'Kitchen Light
            Insteon(21).DeviceName = "Kitchen Light"
            Insteon(21).DeviceDimmable = False
            Insteon(22).DeviceID = "22.85.9B"  'Play Room East (L)
            Insteon(22).DeviceName = "Play Room Light East (L)"
            Insteon(22).DeviceDimmable = False
            Insteon(23).DeviceID = "2B.80.FF"  'Front Hall
            Insteon(23).DeviceName = "Front Hall"
            Insteon(23).DeviceDimmable = False
            Insteon(24).DeviceID = "29.2B.1F"  'Kitchen Nook
            Insteon(24).DeviceName = "Kitchen Nook"
            Insteon(24).DeviceDimmable = False
            Insteon(25).DeviceID = "22.10.A1"  'Play Room West (R)
            Insteon(25).DeviceName = "Play Room Light West (R)"
            Insteon(25).DeviceDimmable = False
            Insteon(26).DeviceID = "2B.80.6A"  'Dining Room
            Insteon(26).DeviceName = "Dining Room"
            Insteon(26).DeviceDimmable = False
            Insteon(27).DeviceID = "33.66.15"  'Garage Inside
            Insteon(27).DeviceName = "Garage Inside"
            Insteon(27).DeviceDimmable = False
            Insteon(28).DeviceID = "15.FD.B8"  'GD Coach 2
            Insteon(28).DeviceName = "GD Coach Lights 2"
            Insteon(28).DeviceDimmable = False

            Insteon(29).DeviceID = "29.2B.68"  'Basement Stairs
            Insteon(29).DeviceName = "Basement Stairs"
            Insteon(29).DeviceDimmable = False
            Insteon(30).DeviceID = "29.2B.11"  'Basement N
            Insteon(30).DeviceName = "Basement N"
            Insteon(30).DeviceDimmable = False
            Insteon(31).DeviceID = "22.10.0C"  'Art Room
            Insteon(31).DeviceName = "Art Room"
            Insteon(31).DeviceDimmable = False
            Insteon(32).DeviceID = "29.27.89"  'Basement Mid
            Insteon(32).DeviceName = "Basement Mid"
            Insteon(32).DeviceDimmable = False
            Insteon(33).DeviceID = "22.12.76"  'Server Room
            Insteon(33).DeviceName = "Server Room"
            Insteon(33).DeviceDimmable = False
            Insteon(34).DeviceID = "2B.83.1F"  'Basement S
            Insteon(34).DeviceName = "Basement S"
            Insteon(34).DeviceDimmable = False
            'Insteon().DeviceID = "1F.4A.5F"  'ApplianceLinc
            'Insteon().DeviceName = "ApplianceLinc"
            'Insteon().DeviceDimmable = False

            '------------

            ExpectedStatus(0) = lblMasterBedroomExpectedStatus
            ExpectedStatus(1) = lblEthanBedroomExpectedStatus
            ExpectedStatus(2) = lblSydneyBedroomExpectedStatus
            ExpectedStatus(3) = lblMasterFanExpectedStatus
            ExpectedStatus(4) = lblEthanFanExpectedStatus
            ExpectedStatus(5) = lblSydneyFanExpectedStatus
            ExpectedStatus(6) = lblMasterBathroomExpectedStatus
            ExpectedStatus(7) = lblSpareBedroomExpectedStatus
            ExpectedStatus(8) = lblKidsHallwayExpectedStatus
            ExpectedStatus(9) = lblMasterClosetExpectedStatus
            ExpectedStatus(10) = lblSpareBedroomFanExpectedStatus
            ExpectedStatus(11) = lblKidsBathroomExpectedStatus
            ExpectedStatus(12) = lblMudRoomExpectedStatus
            ExpectedStatus(13) = lblSpiderLampExpectedStatus
            ExpectedStatus(14) = lblPantryExpectedStatus
            ExpectedStatus(15) = lblSconceLightsExpectedStatus
            ExpectedStatus(16) = lblLivingRoomExpectedStatus
            ExpectedStatus(17) = lblSmartBulbExpectedStatus
            ExpectedStatus(18) = lblFDCoachExpectedStatus
            ExpectedStatus(19) = lblGDCoach1ExpectedStatus
            ExpectedStatus(20) = lblBackyardSpotExpectedStatus
            ExpectedStatus(21) = lblKitchenLightExpectedStatus
            ExpectedStatus(22) = lblPlayRoom1ExpectedStatus
            ExpectedStatus(23) = lblFrontHallExpectedStatus
            ExpectedStatus(24) = lblKitchenNookExpectedStatus
            ExpectedStatus(25) = lblPlayRoom2ExpectedStatus
            ExpectedStatus(26) = lblDiningRoomExpectedStatus
            ExpectedStatus(27) = lblGarageInsideExpectedStatus
            ExpectedStatus(28) = lblGDCoach2ExpectedStatus
            ExpectedStatus(29) = lblBasementStairsExpectedStatus
            ExpectedStatus(30) = lblBasementNExpectedStatus
            ExpectedStatus(31) = lblArtRoomExpectedStatus
            ExpectedStatus(32) = lblBasementMidExpectedStatus
            ExpectedStatus(33) = lblServerRoomExpectedStatus
            ExpectedStatus(34) = lblBasementSExpectedStatus
            'ExpectedStatus() = lblLampExpectedStatus

            ActualStatus(0) = lblMasterBedroomStatus
            ActualStatus(1) = lblEthanBedroomStatus
            ActualStatus(2) = lblSydneyBedroomStatus
            ActualStatus(3) = lblMasterFanStatus
            ActualStatus(4) = lblEthanFanStatus
            ActualStatus(5) = lblSydneyFanStatus
            ActualStatus(6) = lblMasterBathroomStatus
            ActualStatus(7) = lblSpareBedroomStatus
            ActualStatus(8) = lblKidsHallwayStatus
            ActualStatus(9) = lblMasterClosetStatus
            ActualStatus(10) = lblSpareBedroomFanStatus
            ActualStatus(11) = lblKidsBathroomStatus
            ActualStatus(12) = lblMudRoomStatus
            ActualStatus(13) = lblSpiderLampStatus
            ActualStatus(14) = lblPantryStatus
            ActualStatus(15) = lblSconceLightsStatus
            ActualStatus(16) = lblLivingRoomStatus
            ActualStatus(17) = lblSmartBulbStatus
            ActualStatus(18) = lblFDCoachStatus
            ActualStatus(19) = lblGDCoach1Status
            ActualStatus(20) = lblBackyardSpotStatus
            ActualStatus(21) = lblKitchenLightStatus
            ActualStatus(22) = lblPlayRoom1Status
            ActualStatus(23) = lblFrontHallStatus
            ActualStatus(24) = lblKitchenNookStatus
            ActualStatus(25) = lblPlayRoom2Status
            ActualStatus(26) = lblDiningRoomStatus
            ActualStatus(27) = lblGarageInsideStatus
            ActualStatus(28) = lblGDCoach2Status
            ActualStatus(29) = lblBasementStairsStatus
            ActualStatus(30) = lblBasementNStatus
            ActualStatus(31) = lblArtRoomStatus
            ActualStatus(32) = lblBasementMidStatus
            ActualStatus(33) = lblServerRoomStatus
            ActualStatus(34) = lblBasementSStatus
            'ActualStatus() = lblLampStatus

            DisabledStatus(0) = lblMasterBedroomDisabled
            DisabledStatus(1) = lblEthanBedroomDisabled
            DisabledStatus(2) = lblSydneyBedroomDisabled
            DisabledStatus(3) = lblMasterFanDisabled
            DisabledStatus(4) = lblEthanFanDisabled
            DisabledStatus(5) = lblSydneyFanDisabled
            DisabledStatus(6) = lblMasterBathroomDisabled
            DisabledStatus(7) = lblSpareBedroomDisabled
            DisabledStatus(8) = lblKidsHallwayDisabled
            DisabledStatus(9) = lblMasterClosetDisabled
            DisabledStatus(10) = lblSpareBedroomFanDisabled
            DisabledStatus(11) = lblKidsBathroomDisabled
            DisabledStatus(12) = lblMudRoomDisabled
            DisabledStatus(13) = lblSpiderLampDisabled
            DisabledStatus(14) = lblPantryDisabled
            DisabledStatus(15) = lblSconceLightsDisabled
            DisabledStatus(16) = lblLivingRoomDisabled
            DisabledStatus(17) = lblSmartBulbDisabled
            DisabledStatus(18) = lblFDCoachDisabled
            DisabledStatus(19) = lblGDCoach1Disabled
            DisabledStatus(20) = lblBackyardSpotDisabled
            DisabledStatus(21) = lblKitchenLightDisabled
            DisabledStatus(22) = lblPlayRoom1Disabled
            DisabledStatus(23) = lblFrontHallDisabled
            DisabledStatus(24) = lblKitchenNookDisabled
            DisabledStatus(25) = lblPlayRoom2Disabled
            DisabledStatus(26) = lblDiningRoomDisabled
            DisabledStatus(27) = lblGarageInsideDisabled
            DisabledStatus(28) = lblGDCoach2Disabled
            DisabledStatus(29) = lblBasementStairsDisabled
            DisabledStatus(30) = lblBasementNDisabled
            DisabledStatus(31) = lblArtRoomDisabled
            DisabledStatus(32) = lblBasementMidDisabled
            DisabledStatus(33) = lblServerRoomDisabled
            DisabledStatus(34) = lblBasementSDisabled
            'DisabledStatus() = lblLampDisabled

            OnButton(0) = btnMasterBedroomOn
            OnButton(1) = btnEthanBedroomOn
            OnButton(2) = btnSydneyBedroomOn
            OnButton(3) = btnMasterFanOn
            OnButton(4) = btnEthanFanOn
            OnButton(5) = btnSydneyFanOn
            OnButton(6) = btnMasterBathroomOn
            OnButton(7) = btnSpareBedroomOn
            OnButton(8) = btnKidsHallwayOn
            OnButton(9) = btnMasterClosetOn
            OnButton(10) = btnSpareBedroomFanOn
            OnButton(11) = btnKidsBathroomOn
            OnButton(12) = btnMudRoomOn
            OnButton(13) = btnSpiderLampOn
            OnButton(14) = btnPantryOn
            OnButton(15) = btnSconceLightsOn
            OnButton(16) = btnLivingRoomOn
            OnButton(17) = btnSmartBulbOn
            OnButton(18) = btnFDCoachOn
            OnButton(19) = btnGDCoach1On
            OnButton(20) = btnBackyardSpotOn
            OnButton(21) = btnKitchenLightOn
            OnButton(22) = btnPlayRoom1On
            OnButton(23) = btnFrontHallOn
            OnButton(24) = btnKitchenNookOn
            OnButton(25) = btnPlayRoom2On
            OnButton(26) = btnDiningRoomOn
            OnButton(27) = btnGarageInsideOn
            OnButton(28) = btnGDCoach2On
            OnButton(29) = btnBasementStairsOn
            OnButton(30) = btnBasementNOn
            OnButton(31) = btnArtRoomOn
            OnButton(32) = btnBasementMidOn
            OnButton(33) = btnServerRoomOn
            OnButton(34) = btnBasementSOn
            'OnButton() = btnLampOn

            OffButton(0) = btnMasterBedroomOff
            OffButton(1) = btnEthanBedroomOff
            OffButton(2) = btnSydneyBedroomOff
            OffButton(3) = btnMasterFanOff
            OffButton(4) = btnEthanFanOff
            OffButton(5) = btnSydneyFanOff
            OffButton(6) = btnMasterBathroomOff
            OffButton(7) = btnSpareBedroomOff
            OffButton(8) = btnKidsHallwayOff
            OffButton(9) = btnMasterClosetOff
            OffButton(10) = btnSpareBedroomFanOff
            OffButton(11) = btnKidsBathroomOff
            OffButton(12) = btnMudRoomOff
            OffButton(13) = btnSpiderLampOff
            OffButton(14) = btnPantryOff
            OffButton(15) = btnSconceLightsOff
            OffButton(16) = btnLivingRoomOff
            OffButton(17) = btnSmartBulbOff
            OffButton(18) = btnFDCoachOff
            OffButton(19) = btnGDCoach1Off
            OffButton(20) = btnBackyardSpotOff
            OffButton(21) = btnKitchenLightOff
            OffButton(22) = btnPlayRoom1Off
            OffButton(23) = btnFrontHallOff
            OffButton(24) = btnKitchenNookOff
            OffButton(25) = btnPlayRoom2Off
            OffButton(26) = btnDiningRoomOff
            OffButton(27) = btnGarageInsideOff
            OffButton(28) = btnGDCoach2Off
            OffButton(29) = btnBasementStairsOff
            OffButton(30) = btnBasementNOff
            OffButton(31) = btnArtRoomOff
            OffButton(32) = btnBasementMidOff
            OffButton(33) = btnServerRoomOff
            OffButton(34) = btnBasementSOff
            'OffButton() = btnLampOff

            LoadParameters()
            ProgressBar1.Maximum = iRefresh

            If Not PowerLincModemConnect() Then
                Throw New OperationCanceledException("Error Connecting")
            End If

            Me.Refresh()

            AddHandler plm.OnError, AddressOf Me.plm_OnError
            AddHandler plm.SetButton.Tapped, AddressOf Me.plm_SetButton_Tapped
            AddHandler plm.SetButton.PressedAndHeld, AddressOf Me.plm_SetButton_PressedAndHeld
            AddHandler plm.SetButton.ReleasedAfterHolding, AddressOf Me.plm_SetButton_ReleasedAfterHolding
            AddHandler plm.SetButton.UserReset, AddressOf Me.SetButton_UserReset
            AddHandler plm.Network.AllLinkingCompleted, AddressOf Me.Network_AllLinkingCompleted
            AddHandler plm.Network.StandardMessageReceived, AddressOf Me.Network_StandardMessageReceived
            AddHandler plm.Network.X10.UnitAddressed, AddressOf Me.Network_X10_UnitAddressed
            AddHandler plm.Network.X10.CommandReceived, AddressOf Me.Network_X10_CommandReceived
            tEvents.Start()

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Sub InitializeTable()
        Dim bExists As Boolean
        Dim i, j As Integer

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        tNewRequests.Stop()
        tRefreshTable.Stop()

        LineShape1.Visible = False
        LineShape2.Visible = False

        For i = 0 To iNumDevices - 1
            DisabledStatus(i).Text = "_"
        Next i

        For i = 0 To iNumDevices - 1
            If Insteon(i).DeviceActive Then
                Insteon(i).DeviceEnabled = True
                j = 0
                bExists = False
                While bExists = False And j <= iRetries
                    bExists = CheckLampExists(Insteon(i).DeviceID)
                    j += 1
                    Thread.Sleep(iLongDelay)
                End While

                If bStartup Then
                    Try
                        If Insteon_ControlTableAdapter1.Get_Current_State(Insteon(i).DeviceID) >= 1 And Not Insteon(i).DeviceDimmable Then
                            ActualStatus(i).Text = "On"
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                            Insteon(i).DeviceState = 1
                        ElseIf Insteon_ControlTableAdapter1.Get_Current_State(Insteon(i).DeviceID) >= 1 And Insteon(i).DeviceDimmable Then
                            ActualStatus(i).Text = Insteon_ControlTableAdapter1.Get_Current_State(Insteon(i).DeviceID)
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                            Insteon(i).DeviceState = CInt(ActualStatus(i).Text)
                        Else
                            ActualStatus(i).Text = "Off"
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Regular)
                            Insteon(i).DeviceState = 0
                        End If
                    Catch ex As Exception
                    End Try
                End If

                If bExists Then
                    If VerifyLampStatus(Insteon(i).DeviceID) Then
                        OnButton(i).Enabled = True
                        OffButton(i).Enabled = True
                        OnButton(i).BackColor = SystemColors.Control
                        OffButton(i).BackColor = SystemColors.Control
                        DisabledStatus(i).Text = "."
                    Else
                        OnButton(i).Enabled = True
                        OffButton(i).Enabled = True
                        OnButton(i).BackColor = Color.Red
                        OffButton(i).BackColor = Color.Red
                        DisabledStatus(i).Text = "!"
                    End If
                Else
                    OnButton(i).Enabled = True
                    OffButton(i).Enabled = True
                    OnButton(i).BackColor = Color.Red
                    OffButton(i).BackColor = Color.Red
                    DisabledStatus(i).Text = "!"
                End If

                Insteon(i).DeviceEnabled = bExists
                Me.Refresh()
                Thread.Sleep(iShortDelay)
            Else  'Device has been manually disabled
                Insteon(i).DeviceEnabled = False
                OnButton(i).Enabled = False
                OffButton(i).Enabled = False
                OnButton(i).BackColor = Color.Red
                OffButton(i).BackColor = Color.Red
                DisabledStatus(i).Text = "!"
            End If
        Next

        If lblPantryStatus.Text = "On" And lblPantryExpectedStatus.Text = "_" And chkPantryTimer.Checked And lblPantryDisabled.Text = "." Then
            lblPantryDisabled.Text = "c"
            If Not tPantry.Enabled Then
                tPantry.Start()
                addEventMessage("Pantry Timer Started")
            End If
        Else
            If lblPantryDisabled.Text <> "!" Then
                tPantry.Stop()
                lblPantryDisabled.Text = "."
            Else
                tPantry.Stop()
            End If
        End If

        If lblKidsBathroomStatus.Text = "On" And lblKidsBathroomExpectedStatus.Text = "_" And chkKidsBathroomTimer.Checked And lblKidsBathroomDisabled.Text = "." Then
            lblKidsBathroomDisabled.Text = "c"
            If Not tKidsBathroom.Enabled Then
                tKidsBathroom.Start()
                addEventMessage("Kids Bathroom Timer Started")
            End If
        Else
            If lblKidsBathroomDisabled.Text <> "!" Then
                tKidsBathroom.Stop()
                lblKidsBathroomDisabled.Text = "."
            Else
                tKidsBathroom.Stop()
            End If
        End If

        If lblMasterClosetStatus.Text = "On" And lblMasterClosetExpectedStatus.Text = "_" And chkMasterClosetTimer.Checked And lblMasterClosetDisabled.Text = "." Then
            lblMasterClosetDisabled.Text = "c"
            If Not tMasterCloset.Enabled Then
                tMasterCloset.Start()
                addEventMessage("Master Closet Timer Started")
            End If
        Else
            If lblMasterClosetDisabled.Text <> "!" Then
                tMasterCloset.Stop()
                lblMasterClosetDisabled.Text = "."
            Else
                tMasterCloset.Stop()
            End If
        End If

        If lblMudRoomStatus.Text = "On" And lblMudRoomExpectedStatus.Text = "_" And chkMudRoomTimer.Checked And lblMudRoomDisabled.Text = "." Then
            lblMudRoomDisabled.Text = "c"
            If Not tMudRoom.Enabled Then
                tMudRoom.Start()
                addEventMessage("Mud Room Timer Started")
            End If
        Else
            If lblMudRoomDisabled.Text <> "!" Then
                tMudRoom.Stop()
                lblMudRoomDisabled.Text = "."
            Else
                tMudRoom.Stop()
            End If
        End If

        If lblBasementStairsStatus.Text = "On" And lblBasementStairsExpectedStatus.Text = "_" And chkBasementStairsTimer.Checked And lblBasementStairsDisabled.Text = "." Then
            lblBasementStairsDisabled.Text = "c"
            If Not tBasementStairs.Enabled Then
                tBasementStairs.Start()
                addEventMessage("Basement Stairs Timer Started")
            End If
        Else
            If lblBasementStairsDisabled.Text <> "!" Then
                tBasementStairs.Stop()
                lblBasementStairsDisabled.Text = "."
            Else
                tBasementStairs.Stop()
            End If
        End If

        bStartup = False
        bLeakCheckSleep = False
        If frmControls.ckHandleEvents.Checked Then
            tNewRequests.Start()
        End If
        tRefreshTable.Start()
        Me.Cursor = System.Windows.Forms.Cursors.Default
        LineShape1.Visible = True
        LineShape2.Visible = True

    End Sub
#End Region

#Region "Event Handling"
    Private Sub plm_OnError()
        addEventMessage("     " & plm.Exception.Message)
        If plm.Exception.Message = "The port 'COM4' does not exist." Then
            tEvents.Stop()
            tNewRequests.Stop()
            addEventMessage("COM Port does not exist.  Disabling Timers.")
            frmControls.btnConnect.Enabled = True
            frmControls.btnDisconnect.Enabled = False
            frmControls.btnRefresh.Enabled = False
        End If
        If plm.Exception.Message = "Access to the port 'COM4' is denied." Then
            tEvents.Stop()
            tNewRequests.Stop()
            addEventMessage("COM Port access denied.  Disabling Timers.")
            frmControls.btnConnect.Enabled = True
            frmControls.btnDisconnect.Enabled = False
            frmControls.btnRefresh.Enabled = False
        End If
    End Sub

    Private Sub plm_SetButton_Tapped()
        addEventMessage("     PLM SET Button Tapped.")
    End Sub

    Private Sub plm_SetButton_PressedAndHeld()
        addEventMessage("     PLM SET Button Pressed and Held.")
    End Sub

    Private Sub plm_SetButton_ReleasedAfterHolding()
        addEventMessage("     PLM SET Button Released After Holding.")
    End Sub

    Private Sub SetButton_UserReset()
        addEventMessage("     PLM Reset by User.")
    End Sub

    Private Sub Network_AllLinkingCompleted(ByVal sender As Object, ByVal e As AllLinkingCompletedArgs)
        addEventMessage("     All-Linking Completed: " + e.AllLinkingAction.ToString() + ", device: " + e.PeerId.ToString())

    End Sub

    Private Sub Network_StandardMessageReceived(ByVal sender As Object, ByVal e As StandardMessageReceivedArgs)
        Dim command As String = e.Description
        Dim i As Integer
        Dim bFound As Boolean = False

        If (command.Length = 0) Then
            command = "0x" + e.Command1.ToString("X") + ",0x" + e.Command2.ToString("X")
        End If
        addEventMessage("     Standard Message Received: command=" + command + ", description=" +
                    e.Description.ToString() + ", from: " +
                    e.PeerId.ToString())

        If e.MessageType.ToString = "GroupBroadcast" Then  'Prevents multiple triggers
            For i = 0 To iNumDevices - 1
                If Insteon(i).DeviceID = e.PeerId.ToString Then
                    bFound = True
                    Exit For
                End If
            Next
            If Not bFound Then
                Select Case e.PeerId.ToString
                    Case "29.2C.00"  'Kids Hallway 2nd switch
                        i = 8
                        bFound = True
                    Case "29.2B.6C"  'Kitchen Nook 2nd switch
                        i = 24
                        bFound = True
                    Case "29.2A.5F"  'Mud Room 2nd switch
                        i = 12
                        bFound = True
                    Case "29.2B.B7"  'Basement Stairs 2nd switch
                        i = 28
                        bFound = True
                    Case "2B.82.1A"      'Basement S 2nd switch
                        i = 33
                        bFound = True
                    Case "2B.84.0D"      'Dining Room 2nd switch
                        i = 26
                        bFound = True
                    Case "2B.83.AA"      'Front Hall 2nd switch
                        i = 23
                        bFound = True
                    Case "33.65.98"      'Garage Inside 2nd switch
                        i = 27
                        bFound = True
                    Case "31.CB.D4"      'Sump Pump Leak Detect
                        If command = "Turn On to level 2" And ckbLeakCheckActive.Checked And Not bLeakCheckSleep Then    ' Level 1 = Dry Detect;  Level 2 = Wet Detect
                            Event_HistoryTableAdapter1.InsertQuery("5034")
                            Event_Current_StateTableAdapter.AlertLeakDetected()
                            bLeakCheckSleep = True
                        End If
                    Case "2D.61.D8"      'Water Heater Leak Detect
                        If command = "Turn On to level 2" And ckbLeakCheckActive.Checked And Not bLeakCheckSleep Then    ' Level 1 = Dry Detect;  Level 2 = Wet Detect
                            Event_HistoryTableAdapter1.InsertQuery("5034")
                            Event_Current_StateTableAdapter.AlertLeakDetected()
                            bLeakCheckSleep = True
                        End If
                End Select
            End If

            If bFound Then
                Try
                    If command = "Turn On" Then
                        ActualStatus(i).Text = "On"
                        ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                        ExpectedStatus(i).Text = "_"
                        Insteon_ControlTableAdapter1.Current_State_Change(1, e.PeerId.ToString)
                        Insteon(i).DeviceState = 1
                        StartOffTimer(e.PeerId.ToString)   'Not every light has one, but doesn't hurt to send for all since function filters
                        LogEvent(e.PeerId.ToString, 1)
                        If lblLightOnSource.Text = "." Then   'Manual activation of lights
                            lblLightOnSource.Text = "M"
                            Event_Current_StateTableAdapter.UpdateHouseOccupied()
                            Insteon_HistoryTableAdapter1.InsertQuery("1000")  'Manual light activation
                            tLightOnSource.Start()
                        End If
                    End If
                    If command = "End Manual Brightening/Dimming" Then
                        ActualStatus(i).Text = "Dim"
                        ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                        ExpectedStatus(i).Text = "_"
                        If lblLightOnSource.Text = "." Then   'Manual activation of lights
                            lblLightOnSource.Text = "M"
                            Event_Current_StateTableAdapter.UpdateHouseOccupied()
                            Insteon_HistoryTableAdapter1.InsertQuery("1000")  'Manual light activation
                            tLightOnSource.Start()
                        End If
                    End If
                    If command = "Turn Off to level 0" Then
                        ActualStatus(i).Text = "Off"
                        ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Regular)
                        ExpectedStatus(i).Text = "_"
                        Insteon_ControlTableAdapter1.Current_State_Change(0, e.PeerId.ToString)
                        Insteon(i).DeviceState = 0
                        LogEvent(e.PeerId.ToString, 0)
                    End If
                    If command = "Wet Detected" Then

                    End If
                Catch ex As Exception
                End Try
            End If
        End If

    End Sub

    Private Sub Network_X10_UnitAddressed(ByVal sender As Object, ByVal e As X10UnitAddressedArgs)
        addEventMessage("     X10 Unit Addressed: House Code " + e.HouseCode + ", Unit Code: " + e.UnitCode.ToString())
    End Sub

    Private Sub Network_X10_CommandReceived(ByVal sender As Object, ByVal e As X10CommandReceivedArgs)
        addEventMessage("     X10 Command Received: House Code " + e.HouseCode + ", Command: " + e.Command.ToString())
    End Sub


    Private Sub addEventMessage(message As String)

        Dim MESSAGE_QUEUE_LENGTH As Integer = 8

        Dim messageWithTimestamp As String = DateTime.Now.ToString("HH:mm:ss") + " " + message + Chr(13) + Chr(10)
        frmEvents.txtEvents.Text = messageWithTimestamp & frmEvents.txtEvents.Text

    End Sub
#End Region

#Region "Buttons"
    Private Sub btnLampOn_Click(sender As Object, e As EventArgs)
        Try
            If btnLampOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "1F.4A.5F")
                lblLampExpectedStatus.Text = "1"
            Else
                AttemptConnect("1F.4A.5F")
            End If
        Catch ex As Exception
        End Try

    End Sub

    Private Sub btnLampOff_Click(sender As Object, e As EventArgs)
        Try
            If btnLampOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "1F.4A.5F")
                lblLampExpectedStatus.Text = "0"
            Else
                AttemptConnect("1F.4A.5F")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpiderLampOn_Click(sender As Object, e As EventArgs) Handles btnSpiderLampOn.Click
        Try
            If btnSpiderLampOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "1A.F4.47")
                lblSpiderLampExpectedStatus.Text = "1"
            Else
                AttemptConnect("1A.F4.47")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpiderLampOff_Click(sender As Object, e As EventArgs) Handles btnSpiderLampOff.Click
        Try
            If btnSpiderLampOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "1A.F4.47")
                lblSpiderLampExpectedStatus.Text = "0"
            Else
                AttemptConnect("1A.F4.47")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnFDCoachOn_Click(sender As Object, e As EventArgs) Handles btnFDCoachOn.Click
        Try
            If btnFDCoachOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "33.67.22")
                lblFDCoachExpectedStatus.Text = "1"
            Else
                AttemptConnect("33.67.22")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnFDCoachOff_Click(sender As Object, e As EventArgs) Handles btnFDCoachOff.Click
        Try
            If btnFDCoachOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "33.67.22")
                lblFDCoachExpectedStatus.Text = "0"
            Else
                AttemptConnect("33.67.22")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGDCoachOn_Click(sender As Object, e As EventArgs) Handles btnGDCoach1On.Click
        Try
            If btnGDCoach1On.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "17.A6.A9")
                lblGDCoach1ExpectedStatus.Text = "1"
            Else
                AttemptConnect("17.A6.A9")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGDCoachOff_Click(sender As Object, e As EventArgs) Handles btnGDCoach1Off.Click
        Try
            If btnGDCoach1Off.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "17.A6.A9")
                lblGDCoach1ExpectedStatus.Text = "0"
            Else
                AttemptConnect("17.A6.A9")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSmartBulbOn_Click(sender As Object, e As EventArgs) Handles btnSmartBulbOn.Click
        Try
            If btnSmartBulbOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "20.25.3A")
                lblSmartBulbExpectedStatus.Text = "1"
            Else
                AttemptConnect("20.25.3A")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSmartBulbOff_Click(sender As Object, e As EventArgs) Handles btnSmartBulbOff.Click
        Try
            If btnSmartBulbOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "20.25.3A")
                lblSmartBulbExpectedStatus.Text = "0"
            Else
                AttemptConnect("20.25.3A")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnArtRoomOn_Click(sender As Object, e As EventArgs) Handles btnArtRoomOn.Click
        Try
            If btnArtRoomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.10.0C")
                lblArtRoomExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.10.0C")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnArtRoomOff_Click(sender As Object, e As EventArgs) Handles btnArtRoomOff.Click
        Try
            If btnArtRoomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.10.0C")
                lblArtRoomExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.10.0C")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnServerRoomOn_Click(sender As Object, e As EventArgs) Handles btnServerRoomOn.Click
        Try
            If btnServerRoomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.12.76")
                lblServerRoomExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.12.76")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnServerRoomOff_Click(sender As Object, e As EventArgs) Handles btnServerRoomOff.Click
        Try
            If btnServerRoomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.12.76")
                lblServerRoomExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.12.76")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKitchenLightOn_Click(sender As Object, e As EventArgs) Handles btnKitchenLightOn.Click
        Try
            If btnKitchenLightOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.10.09")
                lblKitchenLightExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.10.09")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKitchenLightOff_Click(sender As Object, e As EventArgs) Handles btnKitchenLightOff.Click
        Try
            If btnKitchenLightOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.10.09")
                lblKitchenLightExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.10.09")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnLivingRoomOn_Click(sender As Object, e As EventArgs) Handles btnLivingRoomOn.Click
        Try
            If btnLivingRoomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.13.0E")
                lblLivingRoomExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.13.0E")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnLivingRoomOff_Click(sender As Object, e As EventArgs) Handles btnLivingRoomOff.Click
        Try
            If btnLivingRoomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.13.0E")
                lblLivingRoomExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.13.0E")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnKitchenNookOn_Click(sender As Object, e As EventArgs) Handles btnKitchenNookOn.Click
        Try
            If btnKitchenNookOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.2B.1F")
                lblKitchenNookExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.2B.1F")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKitchenNookOff_Click(sender As Object, e As EventArgs) Handles btnKitchenNookOff.Click
        Try
            If btnKitchenNookOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.2B.1F")
                lblKitchenNookExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.2B.1F")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKidsHallwayOn_Click(sender As Object, e As EventArgs) Handles btnKidsHallwayOn.Click
        Try
            If btnKidsHallwayOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.2A.B9")
                lblKidsHallwayExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.2A.B9")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKidsHallwayOff_Click(sender As Object, e As EventArgs) Handles btnKidsHallwayOff.Click
        Try
            If btnKidsHallwayOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.2A.B9")
                lblKidsHallwayExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.2A.B9")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementNOn_Click(sender As Object, e As EventArgs) Handles btnBasementNOn.Click
        Try
            If btnBasementNOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.2B.11")
                lblBasementNExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.2B.11")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementNOff_Click(sender As Object, e As EventArgs) Handles btnBasementNOff.Click
        Try
            If btnBasementNOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.2B.11")
                lblBasementNExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.2B.11")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementMidOn_Click(sender As Object, e As EventArgs) Handles btnBasementMidOn.Click
        Try
            If btnBasementMidOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.27.89")
                lblBasementMidExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.27.89")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementMidOff_Click(sender As Object, e As EventArgs) Handles btnBasementMidOff.Click
        Try
            If btnBasementMidOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.27.89")
                lblBasementMidExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.27.89")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementSOn_Click(sender As Object, e As EventArgs) Handles btnBasementSOn.Click
        Try
            If btnBasementSOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2B.83.1F")
                lblBasementSExpectedStatus.Text = "1"
            Else
                AttemptConnect("2B.83.1F")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementSOff_Click(sender As Object, e As EventArgs) Handles btnBasementSOff.Click
        Try
            If btnBasementMidOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2B.83.1F")
                lblBasementSExpectedStatus.Text = "0"
            Else
                AttemptConnect("2B.83.1F")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSydneyBedroomOn_Click(sender As Object, e As EventArgs) Handles btnSydneyBedroomOn.Click
        Try
            If btnSydneyBedroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(255, "23.C5.30")
                lblSydneyBedroomExpectedStatus.Text = "255"
            Else
                AttemptConnect("23.C5.30")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSydneyBedroomOff_Click(sender As Object, e As EventArgs) Handles btnSydneyBedroomOff.Click
        Try
            If btnSydneyBedroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "23.C5.30")
                lblSydneyBedroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("23.C5.30")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Private Sub btnSydneyBedroomDim_Click(sender As Object, e As EventArgs) Handles btnSydneyBedroomDim.Click
        Try
            If CInt(mtxtSydneyRoomDim.Text) > 255 Then
                mtxtSydneyRoomDim.Text = "255"
            End If
            Insteon_ControlTableAdapter1.Request_State_Change(CInt(mtxtSydneyRoomDim.Text), "23.C5.30")
            lblSydneyBedroomExpectedStatus.Text = mtxtSydneyRoomDim.Text
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEthanBedroomOn_Click(sender As Object, e As EventArgs) Handles btnEthanBedroomOn.Click
        Try
            If btnEthanBedroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(255, "23.B6.EB")
                lblEthanBedroomExpectedStatus.Text = "255"
            Else
                AttemptConnect("23.B6.EB")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEthanBedroomOff_Click(sender As Object, e As EventArgs) Handles btnEthanBedroomOff.Click
        Try
            If btnEthanBedroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "23.B6.EB")
                lblEthanBedroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("23.B6.EB")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEthanBedroomDim_Click(sender As Object, e As EventArgs) Handles btnEthanBedroomDim.Click
        Try
            If CInt(mtxtEthanRoomDim.Text) > 255 Then
                mtxtEthanRoomDim.Text = "255"
            End If
            Insteon_ControlTableAdapter1.Request_State_Change(CInt(mtxtEthanRoomDim.Text), "23.B6.EB")
            lblEthanBedroomExpectedStatus.Text = mtxtEthanRoomDim.Text
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPantryOn_Click(sender As Object, e As EventArgs) Handles btnPantryOn.Click
        Try
            If btnPantryOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.13.0C")
                lblPantryExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.13.0C")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPantryOff_Click(sender As Object, e As EventArgs) Handles btnPantryOff.Click
        Try
            If btnPantryOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.13.0C")
                lblPantryExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.13.0C")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKidsBathroomOn_Click(sender As Object, e As EventArgs) Handles btnKidsBathroomOn.Click
        Try
            If btnKidsBathroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "23.DB.10")
                lblKidsBathroomExpectedStatus.Text = "1"
            Else
                AttemptConnect("23.DB.10")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnKidsBathroomOff_Click(sender As Object, e As EventArgs) Handles btnKidsBathroomOff.Click
        Try
            If btnKidsBathroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "23.DB.10")
                lblKidsBathroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("23.DB.10")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterBedroomOn_Click(sender As Object, e As EventArgs) Handles btnMasterBedroomOn.Click
        Try
            If btnMasterBedroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(255, "20.E5.84")
                lblMasterBedroomExpectedStatus.Text = "255"
            Else
                AttemptConnect("20.E5.84")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterBedroomOff_Click(sender As Object, e As EventArgs) Handles btnMasterBedroomOff.Click
        Try
            If btnMasterBedroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "20.E5.84")
                lblMasterBedroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("20.E5.84")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterBedroomDim_Click(sender As Object, e As EventArgs) Handles btnMasterBedroomDim.Click
        Try
            If CInt(mtxtMasterBedroomDim.Text) > 255 Then
                mtxtMasterBedroomDim.Text = "255"
            End If
            Insteon_ControlTableAdapter1.Request_State_Change(CInt(mtxtMasterBedroomDim.Text), "20.E5.84")
            lblMasterBedroomExpectedStatus.Text = mtxtMasterBedroomDim.Text
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterBathroomOn_Click(sender As Object, e As EventArgs) Handles btnMasterBathroomOn.Click
        Try
            If btnMasterBathroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "23.E2.E4")
                lblMasterBathroomExpectedStatus.Text = "1"
            Else
                AttemptConnect("23.E2.E4")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterBathroomOff_Click(sender As Object, e As EventArgs) Handles btnMasterBathroomOff.Click
        Try
            If btnMasterBathroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "23.E2.E4")
                lblMasterBathroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("23.E2.E4")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterClosetOn_Click(sender As Object, e As EventArgs) Handles btnMasterClosetOn.Click
        Try
            If btnMasterClosetOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "23.EA.1D")
                lblMasterClosetExpectedStatus.Text = "1"
            Else
                AttemptConnect("23.EA.1D")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterClosetOff_Click(sender As Object, e As EventArgs) Handles btnMasterClosetOff.Click
        Try
            If btnMasterClosetOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "23.EA.1D")
                lblMasterClosetExpectedStatus.Text = "0"
            Else
                AttemptConnect("23.EA.1D")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpareBedroomOn_Click(sender As Object, e As EventArgs) Handles btnSpareBedroomOn.Click
        Try
            If btnSpareBedroomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(255, "22.56.9E")
                lblSpareBedroomExpectedStatus.Text = "255"
            Else
                AttemptConnect("22.56.9E")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpareBedroomOff_Click(sender As Object, e As EventArgs) Handles btnSpareBedroomOff.Click
        Try
            If btnSpareBedroomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.56.9E")
                lblSpareBedroomExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.56.9E")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpareBedroomDim_Click(sender As Object, e As EventArgs) Handles btnSpareBedroomDim.Click
        Try
            If CInt(mtxtSpareBedroomDim.Text) > 255 Then
                mtxtSpareBedroomDim.Text = "255"
            End If
            Insteon_ControlTableAdapter1.Request_State_Change(CInt(mtxtSpareBedroomDim.Text), "22.56.9E")
            lblSpareBedroomExpectedStatus.Text = mtxtSpareBedroomDim.Text
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPlayRoom1On_Click(sender As Object, e As EventArgs) Handles btnPlayRoom1On.Click
        Try
            If btnPlayRoom1On.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.85.9B")
                lblPlayRoom1ExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.85.9B")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPlayRoom1Off_Click(sender As Object, e As EventArgs) Handles btnPlayRoom1Off.Click
        Try
            If btnPlayRoom1Off.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.85.9B")
                lblPlayRoom1ExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.85.9B")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPlayRoom2On_Click(sender As Object, e As EventArgs) Handles btnPlayRoom2On.Click
        Try
            If btnPlayRoom2On.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "22.10.A1")
                lblPlayRoom2ExpectedStatus.Text = "1"
            Else
                AttemptConnect("22.10.A1")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnPlayRoom2Off_Click(sender As Object, e As EventArgs) Handles btnPlayRoom2Off.Click
        Try
            If btnPlayRoom2Off.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "22.10.A1")
                lblPlayRoom2ExpectedStatus.Text = "0"
            Else
                AttemptConnect("22.10.A1")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMudRoomOn_Click(sender As Object, e As EventArgs) Handles btnMudRoomOn.Click
        Try
            If btnMudRoomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.27.CE")
                lblMudRoomExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.27.CE")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMudRoomOff_Click(sender As Object, e As EventArgs) Handles btnMudRoomOff.Click
        Try
            If btnMudRoomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.27.CE")
                lblMudRoomExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.27.CE")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSconceLightsOn_Click(sender As Object, e As EventArgs) Handles btnSconceLightsOn.Click
        Try
            If btnSconceLightsOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.27.9D")
                lblSconceLightsExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.27.9D")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSconceLightsOff_Click(sender As Object, e As EventArgs) Handles btnSconceLightsOff.Click
        Try
            If btnSconceLightsOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.27.9D")
                lblSconceLightsExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.27.9D")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementStairsOn_Click(sender As Object, e As EventArgs) Handles btnBasementStairsOn.Click
        Try
            If btnBasementStairsOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "29.2B.68")
                lblBasementStairsExpectedStatus.Text = "1"
            Else
                AttemptConnect("29.2B.68")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBasementStairsOff_Click(sender As Object, e As EventArgs) Handles btnBasementStairsOff.Click
        Try
            If btnBasementStairsOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "29.2B.68")
                lblBasementStairsExpectedStatus.Text = "0"
            Else
                AttemptConnect("29.2B.68")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnRing_Click(sender As Object, e As EventArgs) Handles btnRing.Click
        Try
            Insteon_ControlTableAdapter1.Request_State_Change(1, "G5")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterFanOn_Click(sender As Object, e As EventArgs) Handles btnMasterFanOn.Click
        Try
            If btnMasterFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2A.39.F6")
                lblMasterFanExpectedStatus.Text = "1"
            Else
                AttemptConnect("2A.39.F6")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnMasterFanOff_Click(sender As Object, e As EventArgs) Handles btnMasterFanOff.Click
        Try
            If btnMasterFanOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2A.39.F6")
                lblMasterFanExpectedStatus.Text = "0"
            Else
                AttemptConnect("2A.39.F6")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEthanFanOn_Click(sender As Object, e As EventArgs) Handles btnEthanFanOn.Click
        Try
            If btnEthanFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2A.3B.F0")
                lblEthanFanExpectedStatus.Text = "1"
            Else
                AttemptConnect("2A.3B.F0")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEthanFanOff_Click(sender As Object, e As EventArgs) Handles btnEthanFanOff.Click
        Try
            If btnEthanFanOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2A.3B.F0")
                lblEthanFanExpectedStatus.Text = "0"
            Else
                AttemptConnect("2A.3B.F0")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSydneyFanOn_Click(sender As Object, e As EventArgs) Handles btnSydneyFanOn.Click
        Try
            If btnSydneyFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2A.3D.80")
                lblSydneyFanExpectedStatus.Text = "1"
            Else
                AttemptConnect("2A.3D.80")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSydneyFanOff_Click(sender As Object, e As EventArgs) Handles btnSydneyFanOff.Click
        Try
            If btnSydneyFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2A.3D.80")
                lblSydneyFanExpectedStatus.Text = "0"
            Else
                AttemptConnect("2A.3D.80")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpareBedroomFanOn_Click(sender As Object, e As EventArgs) Handles btnSpareBedroomFanOn.Click
        Try
            If btnSpareBedroomFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2A.3F.41")
                lblSpareBedroomFanExpectedStatus.Text = "1"
            Else
                AttemptConnect("2A.3F.41")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnSpareBedroomFanOff_Click(sender As Object, e As EventArgs) Handles btnSpareBedroomFanOff.Click
        Try
            If btnSpareBedroomFanOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2A.3F.41")
                lblSpareBedroomFanExpectedStatus.Text = "0"
            Else
                AttemptConnect("2A.3F.41")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBackyardSpotOn_Click(sender As Object, e As EventArgs) Handles btnBackyardSpotOn.Click
        Try
            If btnBackyardSpotOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2B.82.42")
                lblBackyardSpotExpectedStatus.Text = "1"
            Else
                AttemptConnect("2B.82.42")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnBackyardSpotOff_Click(sender As Object, e As EventArgs) Handles btnBackyardSpotOff.Click
        Try
            If btnBackyardSpotOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2B.82.42")
                lblBackyardSpotExpectedStatus.Text = "0"
            Else
                AttemptConnect("2B.82.42")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnFrontHallOn_Click(sender As Object, e As EventArgs) Handles btnFrontHallOn.Click
        Try
            If btnFrontHallOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2B.80.FF")
                lblFrontHallExpectedStatus.Text = "1"
            Else
                AttemptConnect("2B.80.FF")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnFrontHallOff_Click(sender As Object, e As EventArgs) Handles btnFrontHallOff.Click
        Try
            If btnFrontHallOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2B.80.FF")
                lblFrontHallExpectedStatus.Text = "0"
            Else
                AttemptConnect("2B.80.FF")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnDiningRoomOn_Click(sender As Object, e As EventArgs) Handles btnDiningRoomOn.Click
        Try
            If btnDiningRoomOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "2B.80.6A")
                lblDiningRoomExpectedStatus.Text = "1"
            Else
                AttemptConnect("2B.80.6A")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnDiningRoomOff_Click(sender As Object, e As EventArgs) Handles btnDiningRoomOff.Click
        Try
            If btnDiningRoomOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "2B.80.6A")
                lblDiningRoomExpectedStatus.Text = "0"
            Else
                AttemptConnect("2B.80.6A")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGarageInsideOn_Click(sender As Object, e As EventArgs) Handles btnGarageInsideOn.Click
        Try
            If btnGarageInsideOn.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "33.66.15")
                lblGarageInsideExpectedStatus.Text = "1"
            Else
                AttemptConnect("33.66.15")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGarageInsideOff_Click(sender As Object, e As EventArgs) Handles btnGarageInsideOff.Click
        Try
            If btnGarageInsideOff.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "33.66.15")
                lblGarageInsideExpectedStatus.Text = "0"
            Else
                AttemptConnect("33.66.15")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGDCoach2On_Click(sender As Object, e As EventArgs) Handles btnGDCoach2On.Click
        Try
            If btnGDCoach2On.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(1, "15.FD.B8")
                lblGDCoach2ExpectedStatus.Text = "1"
            Else
                AttemptConnect("15.FD.B8")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnGDCoach2Off_Click(sender As Object, e As EventArgs) Handles btnGDCoach2Off.Click
        Try
            If btnGDCoach2Off.BackColor = SystemColors.Control Then
                Insteon_ControlTableAdapter1.Request_State_Change(0, "15.FD.B8")
                lblGDCoach2ExpectedStatus.Text = "0"
            Else
                AttemptConnect("15.FD.B8")
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub btnEvents_Click(sender As Object, e As EventArgs) Handles btnEvents.Click
        frmEvents.ShowDialog()
    End Sub

    Private Sub btnControls_Click(sender As Object, e As EventArgs) Handles btnControls.Click
        frmControls.ShowDialog()
    End Sub

#End Region

#Region "Insteon"
    Private Sub AttemptConnect(strDeviceID As String)
        Dim oDevice As DeviceBase = Nothing
        Try
            If plm.Network.TryConnectToDevice(strDeviceID, oDevice) Then
                Dim oLightingControl As LightingControl = oDevice
                Thread.Sleep(iShortDelay)
                Dim onLevel As Byte = 0
                If oLightingControl.TryGetOnLevel(onLevel) Then
                    addEventMessage("Light Connect Success")
                Else
                    addEventMessage("Connected, but could not communicate")
                End If
            Else
                addEventMessage("Light Connect Error")
            End If
        Catch ex As Exception
            addEventMessage("Error Connecting: " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9005")
        End Try

    End Sub


    Private Sub X10Handler(strDevice As String)
        Try
            If strDevice.Trim = "G5" Then
                plm.Network.X10.House("G").Unit("5").Command(X10Command.On)
                Insteon_ControlTableAdapter1.Request_State_Reset(strDevice)
                addEventMessage("X10 Doorbell Rang.")
                LogEvent(strDevice, 1)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Function VerifyLampStatus(strDevice As String) As Boolean
        Dim i, iCounter As Integer

        Try
            For i = 0 To iNumDevices - 1
                If Insteon(i).DeviceID = strDevice Then Exit For
            Next

            If Insteon(i).DeviceEnabled Then
                iCounter = 0
                If ExpectedStatus(i).Text <> "_" Then  'Ready to be Verified
                    If Insteon(i).DeviceDimmable Then
                        While (ActualStatus(i).Text <> ExpectedStatus(i).Text)
                            Me.Refresh()
                            addEventMessage("Retry.")
                            If InsteonLampChange(strDevice, CInt(ExpectedStatus(i).Text)) Then
                                Insteon(i).DeviceEnabled = True
                                Exit While
                            Else
                                iCounter += 1
                            End If
                            If iCounter >= iRetries Then
                                Insteon(i).DeviceEnabled = False
                                Exit While
                            End If
                            Thread.Sleep(iLongDelay)
                        End While
                    Else  'Non Dimmable
                        While (ActualStatus(i).Text = "On" And CInt(ExpectedStatus(i).Text) = 0)
                            Me.Refresh()
                            addEventMessage("Retry.")
                            If InsteonLampChange(strDevice, 0) Then
                                Insteon(i).DeviceEnabled = True
                                Exit While
                            Else
                                iCounter += 1
                            End If
                            If iCounter >= iRetries Then
                                Insteon(i).DeviceEnabled = False
                                Exit While
                            End If
                            Thread.Sleep(iLongDelay)
                        End While

                        iCounter = 0
                        If ExpectedStatus(i).Text <> "_" Then   'Logic above may turn this value back to "_", so need to recheck
                            While (ActualStatus(i).Text = "Off" And CInt(ExpectedStatus(i).Text) >= 1)
                                Me.Refresh()
                                addEventMessage("Retry.")
                                If InsteonLampChange(strDevice, CInt(ExpectedStatus(i).Text)) Then
                                    Insteon(i).DeviceEnabled = True
                                    Exit While
                                Else
                                    iCounter += 1
                                End If
                                If iCounter >= iRetries Then
                                    Insteon(i).DeviceEnabled = False
                                    Exit While
                                End If
                                Thread.Sleep(iLongDelay)
                            End While
                        End If
                    End If
                        If Insteon(i).DeviceEnabled Then
                        DisabledStatus(i).Text = "."
                        OnButton(i).Enabled = True
                        OffButton(i).Enabled = True
                        OnButton(i).BackColor = SystemColors.Control
                        OffButton(i).BackColor = SystemColors.Control
                    Else
                        DisabledStatus(i).Text = "!"
                        OnButton(i).Enabled = True
                        OffButton(i).Enabled = True
                        OnButton(i).BackColor = Color.Red
                        OffButton(i).BackColor = Color.Red
                        addEventMessage("Lamp Disabled.")
                    End If
                End If
            End If
            Return Insteon(i).DeviceEnabled
        Catch ex As Exception
            addEventMessage("Error Verifying Change:  " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9006")
            Insteon(i).DeviceEnabled = False
            Return False
        End Try

    End Function

    Private Sub StartOffTimer(strDevice As String)
        Select Case strDevice.Trim
            Case "22.13.0C"  'Pantry
                If chkPantryTimer.Checked Then
                    lblPantryDisabled.Text = "c"
                    tPantry.Start()
                    addEventMessage("Pantry Timer Started")
                End If
            Case "23.DB.10"  'Kids Bathroom
                If chkKidsBathroomTimer.Checked Then
                    lblKidsBathroomDisabled.Text = "c"
                    tKidsBathroom.Start()
                    addEventMessage("Kids Bathroom Timer Started")
                End If
            Case "23.EA.1D"  'Master Closet
                If chkMasterClosetTimer.Checked Then
                    lblMasterClosetDisabled.Text = "c"
                    tMasterCloset.Start()
                    addEventMessage("Master Closet Timer Started")
                End If
            Case "29.27.CE", "29.2A.5F"  'Mud Room
                If chkMudRoomTimer.Checked Then
                    lblMudRoomDisabled.Text = "c"
                    tMudRoom.Start()
                    addEventMessage("Mud Room Timer Started")
                End If
            Case "29.2B.68", "29.2B.B7"  'Basement Stairs
                If chkBasementStairsTimer.Checked Then
                    lblBasementStairsDisabled.Text = "c"
                    tBasementStairs.Start()
                    addEventMessage("Basement Stairs Timer Started")
                End If
        End Select
    End Sub

    Private Function InsteonLampChange(strDevice As String, iState As Integer) As Boolean
        Dim oDevice As DeviceBase
        Dim i, iRetryVerify As Integer
        Dim bVerifySuccess As Boolean = False

        For i = 0 To iNumDevices - 1
            If Insteon(i).DeviceID = strDevice Then Exit For
        Next

        If (ActualStatus(i).Text = ExpectedStatus(i).Text) Or (ActualStatus(i).Text = "Off" And ExpectedStatus(i).Text = "0") Then
            ExpectedStatus(i).Text = "_"
            addEventMessage("Skipping Change.  Light is already in expected state." & " (" & Insteon(i).DeviceID & ")")
            Try
                Insteon_ControlTableAdapter1.Request_State_Reset(strDevice)
            Catch ex As Exception
            End Try
            Return True
            Exit Function
        End If

        Try
            oDevice = Nothing
            If (plm.Network.TryConnectToDevice(strDevice, oDevice)) Then
                Dim oLightingControl As LightingControl = oDevice
                iRetryVerify = 0
                If iState = 0 Then  'Turn off
                    If oLightingControl.TurnOff() Then
                        While Not bVerifySuccess And iRetryVerify < iRetries
                            Thread.Sleep(iLongDelay)
                            bVerifySuccess = GetLampStatus(strDevice, oLightingControl)
                            iRetryVerify += 1
                        End While
                        If bVerifySuccess Then addEventMessage("Light Off.")
                    End If
                Else  'Turn on
                    If TypeOf oDevice Is DimmableLightingControl And iState > 1 Then
                        If DirectCast(oDevice, DimmableLightingControl).RampOn(iState) Then
                            While Not bVerifySuccess And iRetryVerify < iRetries
                                Thread.Sleep(iLongDelay)
                                bVerifySuccess = GetLampStatus(strDevice, oLightingControl)
                                iRetryVerify += 1
                            End While
                            If bVerifySuccess Then
                                addEventMessage("Light brightness set to " & iState.ToString & ".")
                                StartOffTimer(strDevice)
                            End If
                        End If
                    Else
                        If oLightingControl.TurnOn() Then
                            While Not bVerifySuccess And iRetryVerify < iRetries
                                Thread.Sleep(iLongDelay)
                                bVerifySuccess = GetLampStatus(strDevice, oLightingControl)
                                iRetryVerify += 1
                            End While
                            If bVerifySuccess Then
                                addEventMessage("Light On.")
                                StartOffTimer(strDevice)
                            End If
                        End If
                    End If
                End If
                oLightingControl = Nothing
            Else
                addEventMessage("Light State Connect Error: " & Insteon(i).DeviceName & " (" & Insteon(i).DeviceID & ")")
                Return False
            End If
            Return bVerifySuccess

        Catch ex As Exception
            addEventMessage("Light State Change Error: " & Insteon(i).DeviceName & " (" & Insteon(i).DeviceID & "): " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9007")
            Return False
        End Try

    End Function

    Private Function CheckLampExists(strDevice As String) As Boolean
        Dim oDevice As DeviceBase
        Dim i As Integer
        Dim bSuccess As Boolean = False
        Dim onLevel As Byte = 0

        For i = 0 To iNumDevices - 1
            If Insteon(i).DeviceID = strDevice Then Exit For
        Next

        Try
            oDevice = Nothing
            If bStartup Then   'First time quick check
                bSuccess = plm.Network.TryConnectToDevice(strDevice, oDevice)
            Else
                If plm.Network.TryConnectToDevice(strDevice, oDevice) Then
                    Dim oLightingControl As LightingControl = oDevice
                    Thread.Sleep(iShortDelay)
                    bSuccess = oLightingControl.TryGetOnLevel(onLevel)
                    'Check to see if light state has been manually changed
                    If bSuccess Then
                        If onLevel >= 1 And Not Insteon(i).DeviceDimmable Then
                            Insteon(i).DeviceState = 1
                            ActualStatus(i).Text = "On"
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                        ElseIf onLevel >= 1 And Insteon(i).DeviceDimmable Then
                            Insteon(i).DeviceState = CInt(onLevel)
                            ActualStatus(i).Text = onLevel.ToString
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                        Else
                            Insteon(i).DeviceState = 0
                            ActualStatus(i).Text = "Off"
                            ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Regular)
                        End If
                        Try
                            Insteon_ControlTableAdapter1.Current_State_Change(Insteon(i).DeviceState, strDevice)
                        Catch ex As Exception
                        End Try
                    End If
                End If
            End If
            Return bSuccess

        Catch ex As Exception
            addEventMessage("Check Lamp Exists Error: " & Insteon(i).DeviceName & " (" & Insteon(i).DeviceID & ")  " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9008")
            Return False
        End Try

    End Function

    Private Function GetLampStatus(strDevice As String, ByRef oLightingControl As LightingControl) As Boolean
        Dim i, j As Integer
        Dim bSuccess As Boolean = False
        Dim onLevel As Byte

        For i = 0 To iNumDevices - 1
            If Insteon(i).DeviceID = strDevice Then Exit For
        Next

        Try
            While Not bSuccess And j <= iRetries
                bSuccess = oLightingControl.TryGetOnLevel(onLevel)
                j += 1
                Thread.Sleep(iShortDelay)
            End While
            If bSuccess Then
                If (onLevel >= 1 And Not Insteon(i).DeviceDimmable) Then
                    ActualStatus(i).Text = "On"
                    ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                    Insteon_ControlTableAdapter1.Current_State_Change(1, strDevice)
                    Insteon(i).DeviceState = 1
                ElseIf (onLevel >= 1 And Insteon(i).DeviceDimmable) Then  'Dim
                    ActualStatus(i).Text = CInt(onLevel).ToString
                    ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Bold)
                    Insteon_ControlTableAdapter1.Current_State_Change(CInt(onLevel), strDevice)
                    Insteon(i).DeviceState = CInt(onLevel)
                Else
                    ActualStatus(i).Text = "Off"
                    ActualStatus(i).Font = New Font(ActualStatus(i).Font, FontStyle.Regular)
                    Insteon_ControlTableAdapter1.Current_State_Change(0, strDevice)
                    Insteon(i).DeviceState = 0
                End If
                LogEvent(strDevice, Insteon(i).DeviceState)
                Try
                    Insteon_ControlTableAdapter1.Request_State_Reset(strDevice)
                Catch ex As Exception
                End Try
                ExpectedStatus(i).Text = "_"
                Return True
            Else
                addEventMessage("Device Connection Error: " & Insteon(i).DeviceName & " (" & Insteon(i).DeviceID & ")")
                Insteon(i).DeviceEnabled = False
                Return False
            End If

        Catch ex As Exception
            addEventMessage("Get Lamp Status Error: " & Insteon(i).DeviceName & " (" & Insteon(i).DeviceID & ")  " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9004")
            Return False
        End Try

    End Function

    Private Function SubtractTime(ByRef tStart As Date, ByRef tEnd As Date)
        Dim iMin, iSec As Integer
        Dim strMin, strSec As String
        Dim tTemp As Date

        strSec = ""
        strMin = ""

        If DatePart(DateInterval.Day, tStart) = DatePart(DateInterval.Day, tEnd) Then
            If DatePart(DateInterval.Second, tStart) < DatePart(DateInterval.Second, tEnd) Then
                iMin = DatePart(DateInterval.Minute, tEnd) - DatePart(DateInterval.Minute, tStart)
                iSec = DatePart(DateInterval.Second, tEnd) - DatePart(DateInterval.Second, tStart)
                If iSec = 0 Then
                    strSec = "00"
                ElseIf iSec <= 9 Then
                    strSec = "0" & iSec
                Else
                    strSec = iSec.ToString.Trim
                End If
                If iMin <= 9 Then
                    strMin = "0" & iMin
                Else
                    strMin = iMin.ToString.Trim
                End If
            Else
                iMin = DatePart(DateInterval.Minute, tEnd) - DatePart(DateInterval.Minute, tStart) - 1
                iSec = 60 - DatePart(DateInterval.Second, tStart)
                iSec = iSec + DatePart(DateInterval.Second, tEnd)
                If iSec = 0 Then
                    strSec = "00"
                ElseIf iSec <= 9 Then
                    strSec = "0" & iSec
                Else
                    strSec = iSec.ToString.Trim
                End If
                If iMin <= 9 Then
                    strMin = "0" & iMin
                Else
                    strMin = iMin.ToString.Trim
                End If
            End If
        Else
            tTemp = DatePart(DateInterval.Month, tStart).ToString & "/" & DatePart(DateInterval.Day, tStart).ToString & "/" & DatePart(DateInterval.Year, tStart).ToString & " 11:59:59 PM"
            iMin = DatePart(DateInterval.Minute, tTemp) - DatePart(DateInterval.Minute, tStart)
            iSec = DatePart(DateInterval.Second, tTemp) - DatePart(DateInterval.Second, tStart) + 1
            tTemp = DatePart(DateInterval.Month, tEnd).ToString & "/" & DatePart(DateInterval.Day, tEnd).ToString & "/" & DatePart(DateInterval.Year, tEnd).ToString & " 12:00:00 AM"
            iMin = iMin + DatePart(DateInterval.Minute, tEnd) - DatePart(DateInterval.Minute, tTemp)
            iSec = iSec + DatePart(DateInterval.Second, tEnd) - DatePart(DateInterval.Second, tTemp)
            If iSec >= 60 Then
                iMin = iMin + 1
                iSec = iSec - 60
            End If
            If iSec = 0 Then
                strSec = "00"
            ElseIf iSec <= 9 Then
                strSec = "0" & iSec
            Else
                strSec = iSec.ToString.Trim
            End If
            If iMin <= 9 Then
                strMin = "0" & iMin
            Else
                strMin = iMin.ToString.Trim
            End If
        End If

        If iMin >= 3 Then   'Warn cycle is taking too long (average normal cycle time is 1:00)
            Event_HistoryTableAdapter1.InsertQuery("9946")
        End If

        Return strMin.Trim & ":" & strSec.Trim

    End Function

    Private Sub LogEvent(strDevice As String, iState As Integer)

        Try
            Select Case strDevice.Trim
                Case "1F.4A.5F"  'Test Lamp
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5038")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5039")
                    End If
                Case "17.A6.A9"  'GD Coach 1
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5044")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5045")
                    End If
                Case "33.67.22"  'FD Coach
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5042")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5043")
                    End If
                Case "1A.F4.47"  'Spider Lamp
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5010")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5011")
                    End If
                Case "20.25.3A"  'Smart Bulb
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5040")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5041")
                    End If
                Case "22.10.0C"  'Art Room
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5067")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5068")
                    End If
                Case "22.12.76"  'Server Room
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5069")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5070")
                    End If
                Case "22.10.09"  'Kitchen
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5079")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5080")
                    End If
                Case "22.13.0E"  'Living Room
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5081")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5082")
                    End If
                Case "29.2B.1F"  'Kitchen Nook
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5102")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5103")
                    End If
                Case "29.2A.B9"  'Kids Hallway
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5104")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5105")
                    End If
                Case "29.2B.11"  'Basement N
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5110")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5111")
                    End If
                Case "29.27.89"  'Basement Mid
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5112")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5113")
                    End If
                Case "2B.83.1F"  'Basement S
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5114")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5115")
                    End If

                Case "23.C5.30"  'Sydney's Bedroom Light
                    If iState = 255 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5071")
                    ElseIf iState >= 1 And iState <= 254 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5072")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5073")
                    End If
                Case "23.B6.EB"  'Ethan's Bedroom Light
                    If iState = 255 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5074")
                    ElseIf iState >= 1 And iState <= 254 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5075")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5076")
                    End If
                Case "22.13.0C"  'Pantry
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5077")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5078")
                    End If
                Case "23.DB.10"  'Kids Bathroom Light
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5083")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5084")
                    End If
                Case "20.E5.84"  'Master Bedroom Light
                    If iState = 255 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5085")
                    ElseIf iState >= 1 And iState <= 254 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5086")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5087")
                    End If
                Case "23.E2.E4"  'Master Bathroom Light
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5088")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5089")
                    End If
                Case "23.EA.1D"  'Master Closet
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5090")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5091")
                    End If
                Case "22.56.9E"  'Spare Bedroom Light
                    If iState = 255 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5092")
                    ElseIf iState >= 1 And iState <= 254 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5093")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5094")
                    End If
                Case "22.85.9B"  'Play Room East (L)
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5095")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5096")
                    End If
                Case "22.10.A1"  'Play Room West (R)
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5097")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5098")
                    End If
                Case "29.27.CE"  'Mud Room
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5106")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5107")
                    End If
                Case "29.27.9D"  'Sconce Light
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5108")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5109")
                    End If
                Case "29.2B.68"  'Basement Stairs
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5116")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5117")
                    End If
                Case "2A.3F.41"  'Spare Bedroom Fan
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5118")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5119")
                    End If
                Case "2A.39.F6"  'Master Bedroom Fan
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5120")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5121")
                    End If
                Case "2A.3D.80"  'Sydney Bedroom Fan
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5122")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5123")
                    End If
                Case "2A.3B.F0"  'Ethan Bedroom Fan
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5124")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5125")
                    End If
                Case "2B.82.42"  'Backyard Spot
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5126")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5127")
                    End If
                Case "2B.80.6A"  'Dining Room
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5128")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5129")
                    End If
                Case "2B.80.FF"  'Front Hall
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5130")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5131")
                    End If
                Case "33.66.15"  'Garage Inside
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5132")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5133")
                    End If
                Case "15.FD.B8"  'GD Coach 2
                    If iState = 1 Then
                        Insteon_HistoryTableAdapter1.InsertQuery("5134")
                    Else
                        Insteon_HistoryTableAdapter1.InsertQuery("5135")
                    End If
                Case "G5"        'Doorbell
                    Insteon_HistoryTableAdapter1.InsertQuery("5037")
            End Select
        Catch ex As Exception
        End Try
    End Sub

#End Region

#Region "Lost Focus"


    Private Sub mtxtKidsBathroomTimer_LostFocus(sender As Object, e As EventArgs) Handles mtxtKidsBathroomTimer.LostFocus
        Try
            If CInt(mtxtKidsBathroomTimer.Text) < 100 And CInt(mtxtKidsBathroomTimer.Text) > 0 Then
                tKidsBathroom.Interval = CInt(mtxtKidsBathroomTimer.Text) * 1000 * 60
                SaveParameters()
            Else
                mtxtKidsBathroomTimer.Text = tKidsBathroom.Interval.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub

    Private Sub mtxtPantryTimer_LostFocus(sender As Object, e As EventArgs) Handles mtxtPantryTimer.LostFocus
        Try
            If CInt(mtxtPantryTimer.Text) < 100 And CInt(mtxtPantryTimer.Text) > 0 Then
                tPantry.Interval = CInt(mtxtPantryTimer.Text) * 1000 * 60
                SaveParameters()
            Else
                mtxtPantryTimer.Text = tPantry.Interval.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub

    Private Sub mtxtMasterClosetTimer_LostFocus(sender As Object, e As EventArgs) Handles mtxtMasterClosetTimer.LostFocus
        Try
            If CInt(mtxtMasterClosetTimer.Text) < 100 And CInt(mtxtMasterClosetTimer.Text) > 0 Then
                tMasterCloset.Interval = CInt(mtxtMasterClosetTimer.Text) * 1000 * 60
                SaveParameters()
            Else
                mtxtMasterClosetTimer.Text = tMasterCloset.Interval.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub

    Private Sub mtxtBasementStairs_LostFocus(sender As Object, e As EventArgs) Handles mtxtBasementStairsTimer.LostFocus
        Try
            If CInt(mtxtBasementStairsTimer.Text) < 1000 And CInt(mtxtBasementStairsTimer.Text) > 0 Then
                tBasementStairs.Interval = CInt(mtxtBasementStairsTimer.Text) * 1000 * 60
                SaveParameters()
            Else
                mtxtBasementStairsTimer.Text = tBasementStairs.Interval.ToString
            End If
        Catch ex As Exception
            'Do nothing
        End Try
    End Sub


#End Region

#Region "Timers"

    Private Sub tEvents_Tick(sender As Object, e As EventArgs) Handles tEvents.Tick
        'Runs once every 500ms

        Try
            plm.Receive()
        Catch ex As Exception
            tEvents.Stop()
            addEventMessage("Error Receiving - Event Timer Stopped")
            Insteon_HistoryTableAdapter1.InsertQuery("9000")
        End Try
    End Sub

    Private Sub tKidsBathroom_Tick(sender As Object, e As EventArgs) Handles tKidsBathroom.Tick
        Try
            tKidsBathroom.Stop()
            Insteon_ControlTableAdapter1.Request_State_Change(0, "23.DB.10")
            lblKidsBathroomExpectedStatus.Text = "0"
            lblKidsBathroomStatus.Text = "."
            addEventMessage("Kids Bathroom Timer Stopped")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tPantry_Tick(sender As Object, e As EventArgs) Handles tPantry.Tick
        Try
            tPantry.Stop()
            Insteon_ControlTableAdapter1.Request_State_Change(0, "22.13.0C")
            lblPantryExpectedStatus.Text = "0"
            lblPantryDisabled.Text = "."
            addEventMessage("Pantry Timer Stopped")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tMasterCloset_Tick(sender As Object, e As EventArgs) Handles tMasterCloset.Tick
        Try
            tMasterCloset.Stop()
            Insteon_ControlTableAdapter1.Request_State_Change(0, "23.EA.1D")
            lblMasterClosetExpectedStatus.Text = "0"
            lblMasterClosetDisabled.Text = "."
            addEventMessage("Master Closet Timer Stopped")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tMudRoom_Tick(sender As Object, e As EventArgs) Handles tMudRoom.Tick
        Try
            tMudRoom.Stop()
            Insteon_ControlTableAdapter1.Request_State_Change(0, "29.27.CE")
            lblMudRoomExpectedStatus.Text = "0"
            lblMudRoomDisabled.Text = "."
            addEventMessage("Mud Room Timer Stopped")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tBasementStairs_Tick(sender As Object, e As EventArgs) Handles tBasementStairs.Tick
        Try
            tBasementStairs.Stop()
            Insteon_ControlTableAdapter1.Request_State_Change(0, "29.2B.68")
            lblBasementStairsExpectedStatus.Text = "0"
            lblBasementStairsDisabled.Text = "."
            addEventMessage("Basement Stairs Timer Stopped")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub tLightOnSource_Tick(sender As Object, e As EventArgs) Handles tLightOnSource.Tick
        lblLightOnSource.Text = "."
        tLightOnSource.Stop()
    End Sub

    Private Sub tNewRequests_Tick(sender As Object, e As EventArgs) Handles tNewRequests.Tick
        Dim strDeviceID As String
        Dim iRequestState As Integer
        Dim i, j As Integer
        Dim bFound As Boolean

        Try
            tEvents.Start()
            Dim tabNewRequests As WatchdogDataSet.Insteon_ControlDataTable
            tabNewRequests = Insteon_ControlTableAdapter1.Get_New_Requests

            If tabNewRequests.Rows.Count > 0 Then
                For i = 0 To tabNewRequests.Rows.Count - 1
                    bFound = False
                    strDeviceID = tabNewRequests.Rows(i).Item(2)
                    For j = 0 To iNumDevices - 1
                        If Insteon(j).DeviceID = strDeviceID Then
                            bFound = True
                            Exit For 'Found device in list
                        End If
                    Next

                    If bFound And Insteon(j).DeviceEnabled And Insteon(j).DeviceActive Then
                        iRequestState = CInt(tabNewRequests.Rows(i).Item(4))
                        ExpectedStatus(j).Text = iRequestState.ToString
                        InsteonLampChange(strDeviceID, iRequestState)
                        If iRequestState = 1 Or iRequestState = 255 Then   'Light On was not manual
                            lblLightOnSource.Text = "A"
                            tLightOnSource.Start()
                        End If
                        Thread.Sleep(iShortDelay)
                        VerifyLampStatus(strDeviceID)
                    End If

                    If Not Insteon(j).DeviceActive Then
                        addEventMessage("Skipping Change.  Light is not active." & " (" & Insteon(i).DeviceID & ")")
                        Try
                            Insteon_ControlTableAdapter1.Request_State_Reset(strDeviceID)
                        Catch ex As Exception
                        End Try
                    End If

                    If strDeviceID.Trim.Length = 2 Then
                        X10Handler(strDeviceID)
                    End If
                Next
            End If
        Catch ex As Exception
            addEventMessage("DB Refresh Error:  " & ex.Message)
            Insteon_HistoryTableAdapter1.InsertQuery("9001")
            tNewRequests.Stop()
            tEvents.Stop()
            frmControls.btnConnect.Enabled = True
            frmControls.btnDisconnect.Enabled = False
        End Try
    End Sub

    Private Sub tRefreshTable_Tick(sender As Object, e As EventArgs) Handles tRefreshTable.Tick
        Dim tStart, tEnd As Date
        'Runs every 1 second

        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value = iRefresh Then
            LineShape1.Visible = False
            LineShape2.Visible = False
            tStart = Now
            InitializeTable()
            ProgressBar1.Value = 0
            If frmEvents.txtEvents.Text.Length > 2000 Then
                frmEvents.txtEvents.Text = frmEvents.txtEvents.Text.Substring(1, 2000)
            End If
            lblLastRefresh.Text = Now.ToString
            tEnd = Now
            lblRefreshTime.Text = SubtractTime(tStart, tEnd)
            LineShape1.Visible = True
            LineShape2.Visible = True
        End If

    End Sub

#End Region

#Region "Disable"
    Private Sub lblMasterBedroom_Click(sender As Object, e As EventArgs) Handles lblMasterBedroom.Click
        If Insteon(0).DeviceActive = True Then
            lblMasterBedroom.ForeColor = Color.Red
            Insteon(0).DeviceActive = False
        Else
            lblMasterBedroom.ForeColor = SystemColors.ControlText
            Insteon(0).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblEthanBedroom_Click(sender As Object, e As EventArgs) Handles lblEthanBedroom.Click
        If Insteon(1).DeviceActive = True Then
            lblEthanBedroom.ForeColor = Color.Red
            Insteon(1).DeviceActive = False
        Else
            lblEthanBedroom.ForeColor = SystemColors.ControlText
            Insteon(1).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSydneyBedroom_Click(sender As Object, e As EventArgs) Handles lblSydneyBedroom.Click
        If Insteon(2).DeviceActive = True Then
            lblSydneyBedroom.ForeColor = Color.Red
            Insteon(2).DeviceActive = False
        Else
            lblSydneyBedroom.ForeColor = SystemColors.ControlText
            Insteon(2).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblMasterFan_Click(sender As Object, e As EventArgs) Handles lblMasterFan.Click
        If Insteon(3).DeviceActive = True Then
            lblMasterFan.ForeColor = Color.Red
            Insteon(3).DeviceActive = False
        Else
            lblMasterFan.ForeColor = SystemColors.ControlText
            Insteon(3).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblEthanFan_Click(sender As Object, e As EventArgs) Handles lblEthanFan.Click
        If Insteon(4).DeviceActive = True Then
            lblEthanFan.ForeColor = Color.Red
            Insteon(4).DeviceActive = False
        Else
            lblEthanFan.ForeColor = SystemColors.ControlText
            Insteon(4).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSydneyFan_Click(sender As Object, e As EventArgs) Handles lblSydneyFan.Click
        If Insteon(5).DeviceActive = True Then
            lblSydneyFan.ForeColor = Color.Red
            Insteon(5).DeviceActive = False
        Else
            lblSydneyFan.ForeColor = SystemColors.ControlText
            Insteon(5).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblMasterBathroom_Click(sender As Object, e As EventArgs) Handles lblMasterBathroom.Click
        If Insteon(6).DeviceActive = True Then
            lblMasterBathroom.ForeColor = Color.Red
            Insteon(6).DeviceActive = False
        Else
            lblMasterBathroom.ForeColor = SystemColors.ControlText
            Insteon(6).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSpareBedroom_Click(sender As Object, e As EventArgs) Handles lblSpareBedroom.Click
        If Insteon(7).DeviceActive = True Then
            lblSpareBedroom.ForeColor = Color.Red
            Insteon(7).DeviceActive = False
        Else
            lblSpareBedroom.ForeColor = SystemColors.ControlText
            Insteon(7).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblKidsHallway_Click(sender As Object, e As EventArgs) Handles lblKidsHallway.Click
        If Insteon(8).DeviceActive = True Then
            lblKidsHallway.ForeColor = Color.Red
            Insteon(8).DeviceActive = False
        Else
            lblKidsHallway.ForeColor = SystemColors.ControlText
            Insteon(8).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblMasterCloset_Click(sender As Object, e As EventArgs) Handles lblMasterCloset.Click
        If Insteon(9).DeviceActive = True Then
            lblMasterCloset.ForeColor = Color.Red
            Insteon(9).DeviceActive = False
        Else
            lblMasterCloset.ForeColor = SystemColors.ControlText
            Insteon(9).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSpareBedroomFan_Click(sender As Object, e As EventArgs) Handles lblSpareBedroomFan.Click
        If Insteon(10).DeviceActive = True Then
            lblSpareBedroomFan.ForeColor = Color.Red
            Insteon(10).DeviceActive = False
        Else
            lblSpareBedroomFan.ForeColor = SystemColors.ControlText
            Insteon(10).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblKidsBathroom_Click(sender As Object, e As EventArgs) Handles lblKidsBathroom.Click
        If Insteon(11).DeviceActive = True Then
            lblKidsBathroom.ForeColor = Color.Red
            Insteon(11).DeviceActive = False
        Else
            lblKidsBathroom.ForeColor = SystemColors.ControlText
            Insteon(11).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblMudRoom_Click(sender As Object, e As EventArgs) Handles lblMudRoom.Click
        If Insteon(12).DeviceActive = True Then
            lblMudRoom.ForeColor = Color.Red
            Insteon(12).DeviceActive = False
        Else
            lblMudRoom.ForeColor = SystemColors.ControlText
            Insteon(12).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSpiderLamp_Click(sender As Object, e As EventArgs) Handles lblSpiderLamp.Click
        If Insteon(13).DeviceActive = True Then
            lblSpiderLamp.ForeColor = Color.Red
            Insteon(13).DeviceActive = False
        Else
            lblSpiderLamp.ForeColor = SystemColors.ControlText
            Insteon(13).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblPantry_Click(sender As Object, e As EventArgs) Handles lblPantry.Click
        If Insteon(14).DeviceActive = True Then
            lblPantry.ForeColor = Color.Red
            Insteon(14).DeviceActive = False
        Else
            lblPantry.ForeColor = SystemColors.ControlText
            Insteon(14).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSconceLights_Click(sender As Object, e As EventArgs) Handles lblSconceLights.Click
        If Insteon(15).DeviceActive = True Then
            lblSconceLights.ForeColor = Color.Red
            Insteon(15).DeviceActive = False
        Else
            lblSconceLights.ForeColor = SystemColors.ControlText
            Insteon(15).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblLivingRoom_Click(sender As Object, e As EventArgs) Handles lblLivingRoom.Click
        If Insteon(16).DeviceActive = True Then
            lblLivingRoom.ForeColor = Color.Red
            Insteon(16).DeviceActive = False
        Else
            lblLivingRoom.ForeColor = SystemColors.ControlText
            Insteon(16).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblSmartBulb_Click(sender As Object, e As EventArgs) Handles lblSmartBulb.Click
        If Insteon(17).DeviceActive = True Then
            lblSmartBulb.ForeColor = Color.Red
            Insteon(17).DeviceActive = False
        Else
            lblSmartBulb.ForeColor = SystemColors.ControlText
            Insteon(17).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblFDCoach_Click(sender As Object, e As EventArgs) Handles lblFDCoach.Click
        If Insteon(18).DeviceActive = True Then
            lblFDCoach.ForeColor = Color.Red
            Insteon(18).DeviceActive = False
        Else
            lblFDCoach.ForeColor = SystemColors.ControlText
            Insteon(18).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblGDCoach_Click(sender As Object, e As EventArgs) Handles lblGDCoach1.Click
        If Insteon(19).DeviceActive = True Then
            lblGDCoach1.ForeColor = Color.Red
            Insteon(19).DeviceActive = False
        Else
            lblGDCoach1.ForeColor = SystemColors.ControlText
            Insteon(19).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblBackyardSpot_Click(sender As Object, e As EventArgs) Handles lblBackyardSpot.Click
        If Insteon(20).DeviceActive = True Then
            lblBackyardSpot.ForeColor = Color.Red
            Insteon(20).DeviceActive = False
        Else
            lblBackyardSpot.ForeColor = SystemColors.ControlText
            Insteon(20).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblKitchenLight_Click(sender As Object, e As EventArgs) Handles lblKitchenLight.Click
        If Insteon(21).DeviceActive = True Then
            lblKitchenLight.ForeColor = Color.Red
            Insteon(21).DeviceActive = False
        Else
            lblKitchenLight.ForeColor = SystemColors.ControlText
            Insteon(21).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblPlayRoom1_Click(sender As Object, e As EventArgs) Handles lblPlayRoom1.Click
        If Insteon(22).DeviceActive = True Then
            lblPlayRoom1.ForeColor = Color.Red
            Insteon(22).DeviceActive = False
        Else
            lblPlayRoom1.ForeColor = SystemColors.ControlText
            Insteon(22).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblFrontHall_Click(sender As Object, e As EventArgs) Handles lblFrontHall.Click
        If Insteon(23).DeviceActive = True Then
            lblFrontHall.ForeColor = Color.Red
            Insteon(23).DeviceActive = False
        Else
            lblFrontHall.ForeColor = SystemColors.ControlText
            Insteon(23).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblKitchenNook_Click(sender As Object, e As EventArgs) Handles lblKitchenNook.Click
        If Insteon(24).DeviceActive = True Then
            lblKitchenNook.ForeColor = Color.Red
            Insteon(24).DeviceActive = False
        Else
            lblKitchenNook.ForeColor = SystemColors.ControlText
            Insteon(24).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblPlayRoom2_Click(sender As Object, e As EventArgs) Handles lblPlayRoom2.Click
        If Insteon(25).DeviceActive = True Then
            lblPlayRoom2.ForeColor = Color.Red
            Insteon(25).DeviceActive = False
        Else
            lblPlayRoom2.ForeColor = SystemColors.ControlText
            Insteon(25).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblDiningRoom_Click(sender As Object, e As EventArgs) Handles lblDiningRoom.Click
        If Insteon(26).DeviceActive = True Then
            lblDiningRoom.ForeColor = Color.Red
            Insteon(26).DeviceActive = False
        Else
            lblDiningRoom.ForeColor = SystemColors.ControlText
            Insteon(26).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblGarageInside_Click(sender As Object, e As EventArgs) Handles lblGarageInside.Click
        If Insteon(27).DeviceActive = True Then
            lblGarageInside.ForeColor = Color.Red
            Insteon(27).DeviceActive = False
        Else
            lblGarageInside.ForeColor = SystemColors.ControlText
            Insteon(27).DeviceActive = True
        End If
        SaveParameters()
    End Sub


    Private Sub lblGDCoach2_Click(sender As Object, e As EventArgs) Handles lblGDCoach2.Click
        If Insteon(28).DeviceActive = True Then
            lblGarageInside.ForeColor = Color.Red
            Insteon(28).DeviceActive = False
        Else
            lblGarageInside.ForeColor = SystemColors.ControlText
            Insteon(28).DeviceActive = True
        End If
        SaveParameters()
    End Sub

    Private Sub lblBasementStairs_Click(sender As Object, e As EventArgs) Handles lblBasementStairs.Click
        If Insteon(29).DeviceActive = True Then
            lblBasementStairs.ForeColor = Color.Red
            Insteon(29).DeviceActive = False
        Else
            lblBasementStairs.ForeColor = SystemColors.ControlText
            Insteon(29).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblBasementN_Click(sender As Object, e As EventArgs) Handles lblBasementN.Click
        If Insteon(30).DeviceActive = True Then
            lblBasementN.ForeColor = Color.Red
            Insteon(30).DeviceActive = False
        Else
            lblBasementN.ForeColor = SystemColors.ControlText
            Insteon(30).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblArtRoom_Click(sender As Object, e As EventArgs) Handles lblArtRoom.Click
        If Insteon(31).DeviceActive = True Then
            lblArtRoom.ForeColor = Color.Red
            Insteon(31).DeviceActive = False
        Else
            lblArtRoom.ForeColor = SystemColors.ControlText
            Insteon(31).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblBasementMid_Click(sender As Object, e As EventArgs) Handles lblBasementMid.Click
        If Insteon(32).DeviceActive = True Then
            lblBasementMid.ForeColor = Color.Red
            Insteon(32).DeviceActive = False
        Else
            lblBasementMid.ForeColor = SystemColors.ControlText
            Insteon(32).DeviceActive = True
        End If
        SaveParameters()

    End Sub

    Private Sub lblServerRoom_Click(sender As Object, e As EventArgs) Handles lblServerRoom.Click
        If Insteon(33).DeviceActive = True Then
            lblServerRoom.ForeColor = Color.Red
            Insteon(33).DeviceActive = False
        Else
            lblServerRoom.ForeColor = SystemColors.ControlText
            Insteon(33).DeviceActive = True
        End If
        SaveParameters()
    End Sub

    Private Sub lblBasementS_Click(sender As Object, e As EventArgs) Handles lblBasementS.Click
        If Insteon(34).DeviceActive = True Then
            lblBasementS.ForeColor = Color.Red
            Insteon(34).DeviceActive = False
        Else
            lblBasementS.ForeColor = SystemColors.ControlText
            Insteon(34).DeviceActive = True
        End If
        SaveParameters()
    End Sub

    Private Sub lblLamp_Click(sender As Object, e As EventArgs) Handles lblLamp.Click
        If Insteon(35).DeviceActive = True Then
            lblLamp.ForeColor = Color.Red
            Insteon(35).DeviceActive = False
        Else
            lblLamp.ForeColor = SystemColors.ControlText
            Insteon(35).DeviceActive = True
        End If
        SaveParameters()
    End Sub



#End Region

End Class
