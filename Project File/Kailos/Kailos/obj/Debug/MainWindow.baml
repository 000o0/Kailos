   M S B A M L   `   `   `  ����A  =Kailos, Version=1.0.0.0, Culture=neutral, PublicKeyToken=nullclr-namespace:KailosKailos  5   	       Kailos.MainWindow   5   	   V RMaterialDesignThemes.Wpf, Version=2.5.0.1205, Culture=neutral, PublicKeyToken=nullFmaterialDesign1http://materialdesigninxaml.net/winfx/xaml/themes  5   	   \ XPresentationFramework, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35R NWindowsBase, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35W SPresentationCore, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35D 9http://schemas.microsoft.com/winfx/2006/xaml/presentation    5   	   R NSystem.Xaml, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089:x,http://schemas.microsoft.com/winfx/2006/xaml     5   	   @d2http://schemas.microsoft.com/expression/blend/2008     5   	   Jmc;http://schemas.openxmlformats.org/markup-compatibility/2006     5   	   &localclr-namespace:Kailos     5   	     � Title$  Kailos��5	   	   $	��450��6   $	��800��6%    � 
ResizeMode$ NoResize=�61    � WindowStartupLocation$ CenterScreen=�6G    � Icon$ Resources/Icon_Kailos.ico��6l   .��5
      � .+�5   
   #  MaterialDesignThemes.Wpf.Card  $��50��6   
����+��.��5      �� $��Center=�6   $��10,0,0,0q�63   $��25(�6E   $%��./Resources/#Metropolis Regular)�6S   
�������Kailos - ARP Tool6�   5        5   
   $��65��6   $	��300��6*   $
��Left=�66   $��	15,65,0,0q�6Q   .��5      � .+�5      �� $	��Top=�6   $��18(�65   $��5,5,0,0q�6C   $%��./Resources/#Metropolis Regular)�6T   
���   �1. Set Your Network DEVICE6�   �� 5      -   5        $��MyNetworkDeviceBox��5      $��Bottom=�65   $
��Left=�6P   $	��280��6k   $��6,0,0,6q�6w   5      5      5        5   
   $	��129��6   $	��300��6+   $
��Left=�67   $��
15,145,0,0q�6R   .��5      � .+�5      �� $	��Top=�6   $��18(�65   $��5,5,0,0q�6C   $%��./Resources/#Metropolis Regular)�6T   
���   �2. Set Target DEVICE6�   �� 5      -   5        $��TargetIPAddress��5      $	��Top=�61   $
��Left=�6I   $��	10,30,0,0q�6d   $	��280��6w   .��6�   
IP Address6�   �� 5       -   5        $��TargetMACAddress��5       $	��Top=�62   $
��Left=�6J   $��	10,65,0,0q�6e   $	��280��6x   .��6�   MAC Address6�   �� 5!      -   5        $��DownloadIPScanner��5!      $��25��62   $	��186��6>   $��Bottom=�6J   $��Right=�6e   $��0,0,5,5q�6�   $%��./Resources/#Metropolis Regular)�6�   $��18(�6�   .��6�   IP Scanner Download6�   5#      5$        5'   
   $	��110��6   $	��300��6+   $
��Left=�67   $��
15,290,0,0q�6R   .��5(      � .+�5)      �� $	��Top=�6   $��18(�65   $��5,5,0,0q�6C   $%��./Resources/#Metropolis Regular)�6T   
���   �'%3. Set Wi-Fi Gateway IP, Attacker MAC6�   �� 5+      -   5        $��GatewayIPAddress��5+      $	��Top=�62   $
��Left=�6J   $��	10,35,0,0q�6e   $	��280��6x   .��6�   Gateway IP Address6�   �� 5,      -   5        $��AttackerMACAddress��5,      $	��Top=�64   $
��Left=�6L   $��	10,70,0,0q�6g   $	��280��6z   .��6�   Attacker MAC Address6�   5.      5/        51   
   $	��440��6   $	��209��6*   $	��Top=�67   $��Right=�6O   $��	0,65,15,0q�6k   .��52      � .+�53      �� ��54      �� 55       j� Stretch$ UniformToFill=�6%    �� ImageSource$" Resources/BackgroundAst2.png��6=   57      �� 59      $	��Top=�6   $��18(�65   $��5,5,0,0q�6C   $%��./Resources/#Metropolis Regular)�6T   
�������Program Info�� 6�   �� 6�   75This program is a Wifi device paralysis program using�� 6�   ?=ARP vulnerability, and 'Winpcap' must be installed on the PC.68  5:      5;        5=   
   $	��440��6   $	��110��6*   $	��Top=�67   $��Right=�6O   $��
0,290,15,0q�6k   ��5>      �� 5?      $! Resources/BackgroundAst.jpg��6   $ UniformToFill=�6G   .��5B      � .+�5C      �� $	��Top=�6   $��18(�65   $��5,5,0,0q�6C   $%��./Resources/#Metropolis Regular)�6T   
�������4. Start the Attack6�   �� 5D      -   5        $��SwitchButton��5D      $��Bottom=�6-   $��Right=�6H   $��80��6d   $��30��6o   $��	0,0,10,10q�6{   .��6�   Start6�   5E      5F      5H      5I      