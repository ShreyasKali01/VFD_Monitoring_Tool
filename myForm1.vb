Imports System.IO.Ports
Imports Modbus.Device

<DebuggerDisplay("{GetDebuggerDisplay(),nq}")>
Public Class Form1

    Dim Master_Station As ModbusSerialMaster

    Dim Com_port As String
    Dim IF_Connected As Boolean = False

    Dim PB_RESET As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        For Each sp As String In My.Computer.Ports.SerialPortNames
            ComboBox5.Items.Add(sp)
            TB_Timer.Enabled = True
        Next



        Timer1.Stop() 'Read Holding Register 
        Timer2.Stop() 'Read COIL STATUS




        ReadConfig()






    End Sub




    Sub ReadConfig()
        Dim theFile0 As String = Nothing
        info_path0.Text = System.Windows.Forms.Application.StartupPath & "\Saveconfig\Saveconfig.txt"
        theFile0 = info_path0.Text

        Dim ObjReader As New System.IO.StreamReader(info_path0.Text)

        TextBox18.Text = ObjReader.ReadLine
        TextBox12.Text = ObjReader.ReadLine
        TextBox13.Text = ObjReader.ReadLine
        TextBox17.Text = ObjReader.ReadLine
        TextBox14.Text = ObjReader.ReadLine

        TextBox22.Text = ObjReader.ReadLine
        TextBox23.Text = ObjReader.ReadLine
        TextBox24.Text = ObjReader.ReadLine
        Label41.Text = ObjReader.ReadLine  'อ่านค่า Label14.Text ว่าเป็น True หรือ Flase


        ComboBox5.SelectedItem = TextBox18.Text
        ComboBox1.SelectedItem = TextBox12.Text
        ComboBox2.SelectedItem = TextBox13.Text
        ComboBox3.SelectedItem = TextBox14.Text
        ComboBox4.SelectedItem = TextBox17.Text

        ObjReader.Close()






    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        Dim theFile As String = Nothing
        info_path0.Text = System.Windows.Forms.Application.StartupPath & "\Saveconfig\Saveconfig.txt"
        theFile = info_path0.Text

        FileSystem.Kill(info_path0.Text)
        If System.IO.File.Exists(info_path0.Text) Then
            MessageBox.Show("Not Connected")
        Else
            Dim ObjWriter As New System.IO.StreamWriter(info_path0.Text)

            '------------------------------------------------------------------

            ObjWriter.Write(TextBox18.Text + vbCrLf)
            ObjWriter.Write(TextBox12.Text + vbCrLf)
            ObjWriter.Write(TextBox13.Text + vbCrLf)
            ObjWriter.Write(TextBox17.Text + vbCrLf)
            ObjWriter.Write(TextBox14.Text + vbCrLf)

            ObjWriter.Write(TextBox22.Text + vbCrLf)
            ObjWriter.Write(TextBox23.Text + vbCrLf)
            ObjWriter.Write(TextBox24.Text + vbCrLf)
            ObjWriter.Write(Label41.Text + vbCrLf)



            ObjWriter.Close()
            ObjWriter.Dispose()

            ComboBox5.SelectedItem = TextBox18.Text
            ComboBox1.SelectedItem = TextBox12.Text
            ComboBox2.SelectedItem = TextBox13.Text
            ComboBox4.SelectedItem = TextBox17.Text
            ComboBox3.SelectedItem = TextBox14.Text





        End If


    End Sub





    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        TextBox18.Text = ComboBox5.SelectedItem
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        TextBox12.Text = ComboBox1.SelectedItem
    End Sub
    ' Parity
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        TextBox13.Text = ComboBox2.SelectedItem
    End Sub
    ' DataBits
    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        TextBox17.Text = ComboBox4.SelectedItem
    End Sub
    'StopBits
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        TextBox14.Text = ComboBox3.SelectedItem
    End Sub


    '-------------------------- Connect  ------------------------------------
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        read()
    End Sub


    Sub read()


        Try


            If (IF_Connected = False) Then

                IF_Connected = True
                Com_port = ComboBox5.SelectedItem
                SerialPort1.PortName = Com_port
                SerialPort1 = New SerialPort(Com_port, Val(TextBox12.Text), Val(TextBox13.Text).ToString, Val(TextBox17.Text), Val(TextBox14.Text))
                SerialPort1.Open() 'Open 
                Master_Station = ModbusSerialMaster.CreateRtu(SerialPort1)
                Master_Station.Transport.ReadTimeout = 200




                Button2.BackColor = Color.Red
                Button2.Text = "Dis-connect"
                ComboBox5.Enabled = False
                ComboBox1.Enabled = False
                ComboBox2.Enabled = False
                ComboBox3.Enabled = False
                ComboBox4.Enabled = False


                Button6.Visible = True
                Button4.Visible = True

                TB_Timer.Enabled = True




                'ถ้า เชื่อมต่อไม่ได้
            ElseIf (IF_Connected = True) Then
                IF_Connected = False
                SerialPort1.Close() 'Close COM port
                Button2.BackColor = Color.Gray
                Button2.Text = "Connect"
                TB_Timer.Enabled = True
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                ComboBox3.Enabled = True
                ComboBox4.Enabled = True
                ComboBox5.Enabled = True

                Button4.BackColor = Color.White




                Button6.Visible = False
                Button4.Visible = False


                Timer1.Stop()
                Timer2.Stop()

            End If


        Catch ex As Exception


            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

            ' MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub read1_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 200, 13) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox49.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox52.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox53.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox54.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox55.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox56.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox57.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox58.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox59.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox60.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox61.Text = Convert.ToString(holding_register(10)) '16bit 
            TextBox62.Text = Convert.ToString(holding_register(11)) '16bit 
            TextBox63.Text = Convert.ToString(holding_register(12)) '16bit 
            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True





            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub

    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub

    'ทุกครั้งที่ Timer1 ทำงาน ทุกๆ 1 วินาที เพื่อแสดงค่า ที่ Update
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 100, 10) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox28.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox35.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox36.Text = Convert.ToString(holding_register(2))
            TextBox37.Text = Convert.ToString(holding_register(3))
            TextBox38.Text = Convert.ToString(holding_register(4))
            TextBox39.Text = Convert.ToString(holding_register(5))
            TextBox40.Text = Convert.ToString(holding_register(6))
            TextBox41.Text = Convert.ToString(holding_register(7))
            TextBox42.Text = Convert.ToString(holding_register(8))
            TextBox48.Text = Convert.ToString(holding_register(9))
            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True




            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub

    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 0, 14) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox6.Text = Convert.ToString(holding_register(0))   '16bit
            textBox7.Text = Convert.ToString(holding_register(1))   '16bit
            textBox8.Text = Convert.ToString(holding_register(2))   '16bit
            textBox9.Text = Convert.ToString(holding_register(3))  '16bit 
            textBox10.Text = Convert.ToString(holding_register(4))  '16bit
            textBox11.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox15.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox25.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox26.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox50.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox31.Text = Convert.ToString(holding_register(10)) '16bit
            TextBox82.Text = Convert.ToString(holding_register(11)) '16bit
            TextBox83.Text = Convert.ToString(holding_register(12)) '16bit
            TextBox84.Text = Convert.ToString(holding_register(13)) '16bit
            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True



            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button6.BackColor = Color.White
        Timer1.Start()
        Button6.BackColor = Color.Green


    End Sub

    'ปุ่ม หยุดการอ่านค่า  HoldingRegisters
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.BackColor = Color.White
        Timer1.Stop()
        Button4.BackColor = Color.Red



    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Label67_Click(sender As Object, e As EventArgs)

    End Sub

    'ทุกครั้งที่ Timer1 ทำงาน ทุกๆ 1 วินาที เพื่อแสดงค่า ที่ Update
    Private Sub read2_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 300, 9) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox1.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox2.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox3.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox4.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox5.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox16.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox27.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox29.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox30.Text = Convert.ToString(holding_register(8)) '16bit 

            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True




            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub

    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    'ทุกครั้งที่ Timer1 ทำงาน ทุกๆ 1 วินาที เพื่อแสดงค่า ที่ Update
    Private Sub read4_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 500, 21) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox81.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox80.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox79.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox78.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox77.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox76.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox75.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox74.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox73.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox46.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox94.Text = Convert.ToString(holding_register(10)) '16bit 
            TextBox93.Text = Convert.ToString(holding_register(11)) '16bit 
            TextBox91.Text = Convert.ToString(holding_register(12)) '16bit 
            TextBox92.Text = Convert.ToString(holding_register(13)) '16bit 
            TextBox90.Text = Convert.ToString(holding_register(14)) '16bit 
            TextBox89.Text = Convert.ToString(holding_register(15)) '16bit 
            TextBox88.Text = Convert.ToString(holding_register(16)) '16bit 
            TextBox87.Text = Convert.ToString(holding_register(17)) '16bit 
            TextBox86.Text = Convert.ToString(holding_register(18)) '16bit 
            TextBox85.Text = Convert.ToString(holding_register(19)) '16bit 
            TextBox95.Text = Convert.ToString(holding_register(20)) '16bit
            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    'ทุกครั้งที่ Timer1 ทำงาน ทุกๆ 1 วินาที เพื่อแสดงค่า ที่ Update
    Private Sub read3_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 400, 11) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox72.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox71.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox70.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox69.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox68.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox67.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox66.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox65.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox64.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox51.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox47.Text = Convert.ToString(holding_register(10)) '16bit 
            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    'ทุกครั้งที่ Timer1 ทำงาน ทุกๆ 1 วินาที เพื่อแสดงค่า ที่ Update
    Private Sub read5_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 600, 5) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox100.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox99.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox98.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox97.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox96.Text = Convert.ToString(holding_register(4))  '16bit

            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read6_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 700, 10) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox110.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox109.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox108.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox107.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox106.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox105.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox104.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox103.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox102.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox101.Text = Convert.ToString(holding_register(9)) '16bit 

            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read7_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 800, 19) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox120.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox119.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox118.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox117.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox116.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox115.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox114.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox113.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox112.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox111.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox130.Text = Convert.ToString(holding_register(10)) '16bit 
            TextBox129.Text = Convert.ToString(holding_register(11)) '16bit 
            TextBox128.Text = Convert.ToString(holding_register(12)) '16bit 
            TextBox127.Text = Convert.ToString(holding_register(13)) '16bit 
            TextBox126.Text = Convert.ToString(holding_register(14)) '16bit 
            TextBox125.Text = Convert.ToString(holding_register(15)) '16bit 
            TextBox124.Text = Convert.ToString(holding_register(16)) '16bit 
            TextBox123.Text = Convert.ToString(holding_register(17)) '16bit 
            TextBox122.Text = Convert.ToString(holding_register(18)) '16bit 


            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read8_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 900, 33) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox148.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox147.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox146.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox145.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox144.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox143.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox142.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox141.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox138.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox137.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox136.Text = Convert.ToString(holding_register(10)) '16bit 
            TextBox135.Text = Convert.ToString(holding_register(11)) '16bit 
            TextBox134.Text = Convert.ToString(holding_register(12)) '16bit 
            TextBox133.Text = Convert.ToString(holding_register(13)) '16bit 
            TextBox132.Text = Convert.ToString(holding_register(14)) '16bit 
            TextBox131.Text = Convert.ToString(holding_register(15)) '16bit 
            TextBox159.Text = Convert.ToString(holding_register(16)) '16bit 
            TextBox158.Text = Convert.ToString(holding_register(17))   '16bi
            TextBox157.Text = Convert.ToString(holding_register(18))   '16bit
            TextBox156.Text = Convert.ToString(holding_register(19))   '16bit
            TextBox155.Text = Convert.ToString(holding_register(20))   '16bit
            TextBox154.Text = Convert.ToString(holding_register(21))  '16bit 
            TextBox153.Text = Convert.ToString(holding_register(22))  '16bit
            TextBox152.Text = Convert.ToString(holding_register(23))  '16bit 
            TextBox140.Text = Convert.ToString(holding_register(24))  '16bit 
            TextBox139.Text = Convert.ToString(holding_register(25))  '16bit
            TextBox121.Text = Convert.ToString(holding_register(26)) '16bit 
            TextBox149.Text = Convert.ToString(holding_register(27)) '16bit 
            TextBox151.Text = Convert.ToString(holding_register(28)) '16bit 
            TextBox150.Text = Convert.ToString(holding_register(29)) '16bit 
            TextBox161.Text = Convert.ToString(holding_register(30)) '16bit 
            TextBox160.Text = Convert.ToString(holding_register(31)) '16bit 
            TextBox162.Text = Convert.ToString(holding_register(32)) '16bit 



            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read9_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 1100, 7) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox169.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox168.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox167.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox166.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox165.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox164.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox163.Text = Convert.ToString(holding_register(6))  '16bit 




            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub

    Private Sub read10_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 1000, 8) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox177.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox176.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox175.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox174.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox173.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox172.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox171.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox170.Text = Convert.ToString(holding_register(7))  '16bit 



            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub
    Private Sub read11_hold_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        Try

            Dim Slave_ID As Byte = TextBox22.Text
            Dim startAddress As UShort = TextBox23.Text
            Dim numOfPoints As UShort = TextBox24.Text


            TextBox19.Text = Convert.ToString(Slave_ID) '16bit
            TextBox20.Text = Convert.ToString(startAddress - 1) '16bit
            TextBox21.Text = Convert.ToString(numOfPoints) '16bit



            'เชื่อมต่อ เพื่ออ่านค่า ใน Holding Register
            Dim holding_register As UShort() = Master_Station.ReadHoldingRegisters(1, 1200, 17) ' Read holding FC03

            'ให้แสดงค่าลงใน Textbox ตามตำแหน่งของ Address ที่กำหนด  จาก Dim numOfPoints As UShort = 10  ---> 10 คือ ให้อ่านมา 10 Address/10ค่า
            'สร้าง Textbox  ให้รองรับการ แสดงค่าที่อ่านได้จาก Address ที่กำหนดออกมา

            TextBox196.Text = Convert.ToString(holding_register(0))   '16bit
            TextBox195.Text = Convert.ToString(holding_register(1))   '16bit
            TextBox194.Text = Convert.ToString(holding_register(2))   '16bit
            TextBox193.Text = Convert.ToString(holding_register(3))  '16bit 
            TextBox192.Text = Convert.ToString(holding_register(4))  '16bit
            TextBox191.Text = Convert.ToString(holding_register(5))  '16bit 
            TextBox190.Text = Convert.ToString(holding_register(6))  '16bit 
            TextBox189.Text = Convert.ToString(holding_register(7))  '16bit
            TextBox188.Text = Convert.ToString(holding_register(8)) '16bit 
            TextBox187.Text = Convert.ToString(holding_register(9)) '16bit 
            TextBox186.Text = Convert.ToString(holding_register(10)) '16bit 
            TextBox185.Text = Convert.ToString(holding_register(11)) '16bit 
            TextBox184.Text = Convert.ToString(holding_register(12)) '16bit 
            TextBox183.Text = Convert.ToString(holding_register(13)) '16bit 
            TextBox182.Text = Convert.ToString(holding_register(14)) '16bit 
            TextBox181.Text = Convert.ToString(holding_register(15)) '16bit 
            TextBox180.Text = Convert.ToString(holding_register(16)) '16bit


            '--------------------------------------------------------------------------------------


        Catch ex As Exception
            'ถ้ากำหนดค่าผิด ให้Not Connected
            IF_Connected = True
            SerialPort1.Close() 'Close COM port

            Button2.Text = "Connect"
            Button2.BackColor = Color.Gray
            Timer1.Stop()

            TB_Timer.Enabled = True
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            ComboBox3.Enabled = True
            ComboBox4.Enabled = True
            ComboBox5.Enabled = True


            TextBox72.Text = ""
            TextBox71.Text = ""
            TextBox70.Text = ""
            TextBox69.Text = ""
            TextBox68.Text = ""
            TextBox67.Text = ""
            TextBox66.Text = ""
            TextBox65.Text = ""
            TextBox64.Text = ""
            TextBox51.Text = ""
            TextBox47.Text = ""

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Not Connected", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try



    End Sub
    'ปุ่ม สั่งให้อ่านค่า HoldingRegisters ทุกๆ 1 วินาที
    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click

        Timer1.Start() 'ให้  Timer1 ทำงาน โดยตั้งค่า 1 Sec ทำให้มีการอ่านค่าทุกๆ 1 Sec 

    End Sub


    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged

    End Sub

    Private Sub TextBox41_TextChanged(sender As Object, e As EventArgs) Handles TextBox41.TextChanged

    End Sub

    Private Sub TextBox42_TextChanged(sender As Object, e As EventArgs) Handles TextBox42.TextChanged

    End Sub

    Private Sub TextBox64_TextChanged(sender As Object, e As EventArgs) Handles TextBox64.TextChanged

    End Sub

    Private Sub GroupBox7_Enter(sender As Object, e As EventArgs) Handles GroupBox7.Enter

    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged

    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click

    End Sub

    Private Sub TextBox81_TextChanged(sender As Object, e As EventArgs) Handles TextBox81.TextChanged

    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click

    End Sub

    Private Sub GroupBox11_Enter(sender As Object, e As EventArgs) Handles GroupBox11.Enter

    End Sub

    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged

    End Sub

    Private Sub textBox9_TextChanged(sender As Object, e As EventArgs) Handles textBox9.TextChanged

    End Sub

    Private Sub GroupBox5_Enter(sender As Object, e As EventArgs) Handles GroupBox5.Enter

    End Sub

    Private Sub GroupBox8_Enter(sender As Object, e As EventArgs) Handles GroupBox8.Enter

    End Sub

    Private Sub Label126_Click(sender As Object, e As EventArgs) Handles Label126.Click

    End Sub

    Private Sub Label122_Click(sender As Object, e As EventArgs) Handles Label122.Click

    End Sub

    Private Sub Label143_Click(sender As Object, e As EventArgs) Handles Label143.Click

    End Sub

    Private Sub GroupBox12_Enter(sender As Object, e As EventArgs) Handles GroupBox12.Enter

    End Sub

    Private Sub Label166_Click(sender As Object, e As EventArgs) Handles Label166.Click

    End Sub

    Private Sub TextBox55_TextChanged(sender As Object, e As EventArgs) Handles TextBox55.TextChanged

    End Sub

    Private Sub TextBox80_TextChanged(sender As Object, e As EventArgs) Handles TextBox80.TextChanged

    End Sub

    Private Sub TextBox96_TextChanged(sender As Object, e As EventArgs) Handles TextBox96.TextChanged

    End Sub

    Private Sub GroupBox10_Enter(sender As Object, e As EventArgs) Handles GroupBox10.Enter

    End Sub

    Private Sub Label82_Click(sender As Object, e As EventArgs) Handles Label82.Click

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Try
            '-----------------------------------------------------------------------------------

            Dim Slave_ID As Byte = TextBox22.Text
            Dim Data_Input As UShort = TextBox179.Text
            Dim registerAddress As UShort = TextBox178.Text



            Master_Station.WriteSingleRegister(TextBox22.Text, TextBox178.Text, TextBox179.Text)


        Catch ex As Exception

            Dim show As New mb.ShowMessagebox
            show.Fonts(New Font("Arial", 16, FontStyle.Regular))
            show.ShowBox("Invalid Operation", mb.MStyle.ok, mb.FStyle.Information, "Error")

        End Try


    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged

    End Sub

    Private Sub TextBox178_TextChanged(sender As Object, e As EventArgs) Handles TextBox178.TextChanged

    End Sub

    Private Sub Label194_Click(sender As Object, e As EventArgs) Handles Label194.Click

    End Sub

    Private Sub Label204_Click(sender As Object, e As EventArgs) Handles Label204.Click

    End Sub

    Private Sub Label218_Click(sender As Object, e As EventArgs) Handles Label218.Click

    End Sub

    Private Sub Label221_Click(sender As Object, e As EventArgs) Handles Label221.Click

    End Sub

    Private Sub Label225_Click(sender As Object, e As EventArgs) Handles Label225.Click

    End Sub

    Private Sub Label226_Click(sender As Object, e As EventArgs) Handles Label226.Click

    End Sub

    Private Sub Label92_Click(sender As Object, e As EventArgs) Handles Label92.Click

    End Sub

    Private Sub TextBox87_TextChanged(sender As Object, e As EventArgs) Handles TextBox87.TextChanged

    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click

    End Sub

    Private Sub txtFile_TextChanged(sender As Object, e As EventArgs)

    End Sub







    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        TabControl1.SelectedTab = TabPage8
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click
        TabControl1.SelectedTab = TabPage2
    End Sub

    Private Sub GroupBox16_Enter(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40000 ," + TextBox6.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40001 ," + textBox7.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40002 ," + textBox8.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40003 ," + textBox9.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40004 ," + textBox10.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40005 ," + textBox11.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40006 ," + TextBox15.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40007 ," + TextBox25.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40008 ," + TextBox26.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40009 ," + TextBox50.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40010 ," + TextBox31.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40011 ," + TextBox82.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40012 ," + TextBox83.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40013 ," + TextBox84.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40100 ," + TextBox28.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40101 ," + TextBox35.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40102 ," + TextBox36.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40103 ," + TextBox37.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40104 ," + TextBox38.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40105 ," + TextBox39.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40106 ," + TextBox40.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40107 ," + TextBox41.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40108 ," + TextBox42.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40109 ," + TextBox48.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40200 ," + TextBox49.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40201 ," + TextBox52.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40202 ," + TextBox53.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40203 ," + TextBox54.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40204 ," + TextBox55.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40205 ," + TextBox56.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40206 ," + TextBox57.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40207 ," + TextBox58.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40208 ," + TextBox59.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40209 ," + TextBox60.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40210 ," + TextBox61.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40211 ," + TextBox62.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40212 ," + TextBox63.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40300 ," + TextBox1.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40301 ," + TextBox2.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40302 ," + TextBox3.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40303 ," + TextBox4.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40304 ," + TextBox5.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40305 ," + TextBox16.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40306 ," + TextBox27.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40307 ," + TextBox29.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40308 ," + TextBox30.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40400 ," + TextBox72.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40401 ," + TextBox71.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40402 ," + TextBox70.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40403 ," + TextBox69.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40404 ," + TextBox68.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40405 ," + TextBox67.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40406 ," + TextBox66.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40407 ," + TextBox65.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40408 ," + TextBox64.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40409 ," + TextBox51.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40410 ," + TextBox47.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40500 ," + TextBox81.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40501 ," + TextBox80.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40502 ," + TextBox79.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40503 ," + TextBox78.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40504 ," + TextBox77.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40505 ," + TextBox76.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40506 ," + TextBox75.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40507 ," + TextBox74.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40508 ," + TextBox73.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40509 ," + TextBox46.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40510 ," + TextBox94.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40511 ," + TextBox93.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40512 ," + TextBox91.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40513 ," + TextBox92.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40514 ," + TextBox90.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40515 ," + TextBox89.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40516 ," + TextBox88.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40517 ," + TextBox87.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40518 ," + TextBox86.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40519 ," + TextBox85.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40520 ," + TextBox51.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40600 ," + TextBox100.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40601 ," + TextBox99.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40602 ," + TextBox98.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40603 ," + TextBox97.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40604 ," + TextBox96.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40700 ," + TextBox110.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40701 ," + TextBox109.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40702 ," + TextBox108.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40703 ," + TextBox107.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40704 ," + TextBox106.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40705 ," + TextBox105.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40706 ," + TextBox104.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40707 ," + TextBox103.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40708 ," + TextBox102.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40709 ," + TextBox101.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40800 ," + TextBox120.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40801 ," + TextBox119.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40802 ," + TextBox118.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40803 ," + TextBox117.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40804 ," + TextBox116.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40805 ," + TextBox115.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40806 ," + TextBox114.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40807 ," + TextBox113.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40808 ," + TextBox112.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40809 ," + TextBox111.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40810 ," + TextBox130.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40811 ," + TextBox129.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40812 ," + TextBox128.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40813 ," + TextBox127.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40814 ," + TextBox126.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40815 ," + TextBox125.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40816 ," + TextBox124.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40817 ," + TextBox123.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40818 ," + TextBox122.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40900 ," + TextBox148.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40901 ," + TextBox147.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40902 ," + TextBox146.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40903 ," + TextBox145.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40904 ," + TextBox144.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40905 ," + TextBox143.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40906 ," + TextBox142.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40907 ," + TextBox141.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40908 ," + TextBox138.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40909 ," + TextBox137.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40910 ," + TextBox136.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40911 ," + TextBox135.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40912 ," + TextBox134.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40913 ," + TextBox133.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40914 ," + TextBox132.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40915 ," + TextBox131.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40916 ," + TextBox159.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40917 ," + TextBox158.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40918 ," + TextBox157.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40919 ," + TextBox156.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40920 ," + TextBox155.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40921 ," + TextBox154.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40922 ," + TextBox153.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40923 ," + TextBox152.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40924 ," + TextBox140.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40925 ," + TextBox139.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40926 ," + TextBox121.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40927 ," + TextBox149.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40928 ," + TextBox151.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40929 ," + TextBox150.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40930 ," + TextBox161.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40931 ," + TextBox160.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40932 ," + TextBox162.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 40933 ," + TextBox126.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41000 ," + TextBox169.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41001 ," + TextBox168.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41002 ," + TextBox167.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41003 ," + TextBox166.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41004 ," + TextBox165.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41005 ," + TextBox164.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41006 ," + TextBox163.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41100 ," + TextBox177.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41101 ," + TextBox176.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41102 ," + TextBox175.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41103 ," + TextBox174.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41104 ," + TextBox173.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41105 ," + TextBox172.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41106 ," + TextBox171.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41107 ," + TextBox170.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41200 ," + TextBox196.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41201 ," + TextBox195.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41202 ," + TextBox194.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41203 ," + TextBox193.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41204 ," + TextBox192.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41205 ," + TextBox191.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41206 ," + TextBox190.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41207 ," + TextBox189.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41208 ," + TextBox188.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41209 ," + TextBox187.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41210 ," + TextBox186.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41211 ," + TextBox185.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41212 ," + TextBox184.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41213 ," + TextBox183.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41214 ," + TextBox182.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41215 ," + TextBox181.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)
        TextBox197.AppendText(TimeOfDay & vbTab & ", 41216 ," + TextBox180.Text + vbNewLine)
        TextBox197.AppendText(" " + vbNewLine)

        Dim iSave As New SaveFileDialog
        iSave.Filter = " CSV  files (*.csv)  | *.csv "
        iSave.FilterIndex = 2
        iSave.RestoreDirectory = False
        If iSave.ShowDialog() = DialogResult.OK Then
            System.IO.File.WriteAllText(iSave.FileName, TextBox197.Text)
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        TabControl1.SelectedTab = TabPage1
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        TabControl1.SelectedTab = TabPage3
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        TabControl1.SelectedTab = TabPage4
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        TabControl1.SelectedTab = TabPage5
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        TabControl1.SelectedTab = TabPage6
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        TabControl1.SelectedTab = TabPage7
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        TabControl1.SelectedTab = TabPage9
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        TabControl1.SelectedTab = TabPage10
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        TabControl1.SelectedTab = TabPage11
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        TabControl1.SelectedTab = TabPage12
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        TabControl1.SelectedTab = TabPage13
    End Sub

    Private Sub GroupBox6_Enter(sender As Object, e As EventArgs) Handles GroupBox6.Enter

    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click

    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click

    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click

    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click

    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub Label50_Click(sender As Object, e As EventArgs) Handles Label50.Click

    End Sub

    Private Sub Label51_Click(sender As Object, e As EventArgs) Handles Label51.Click

    End Sub

    Private Sub Label106_Click(sender As Object, e As EventArgs) Handles Label106.Click

    End Sub

    Private Sub TextBox122_TextChanged(sender As Object, e As EventArgs) Handles TextBox122.TextChanged

    End Sub

    Private Sub Label115_Click(sender As Object, e As EventArgs) Handles Label115.Click

    End Sub

    Private Sub Label109_Click(sender As Object, e As EventArgs) Handles Label109.Click

    End Sub

    Private Sub Label113_Click(sender As Object, e As EventArgs) Handles Label113.Click

    End Sub

    Private Sub Label116_Click(sender As Object, e As EventArgs) Handles Label116.Click

    End Sub

    Private Sub Label117_Click(sender As Object, e As EventArgs) Handles Label117.Click

    End Sub

    Private Sub Label118_Click(sender As Object, e As EventArgs) Handles Label118.Click

    End Sub

    Private Sub Label294_Click(sender As Object, e As EventArgs) Handles Label294.Click

    End Sub

    Private Sub GroupBox15_Enter(sender As Object, e As EventArgs) Handles GroupBox15.Enter

    End Sub

    Private Sub TabPage14_Click(sender As Object, e As EventArgs) Handles TabPage14.Click

    End Sub
End Class

