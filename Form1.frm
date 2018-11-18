VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "智能室内环境监测系统"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   FillColor       =   &H00C0C0FF&
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   12
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   13200
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   9
      Left            =   4680
      TabIndex        =   21
      Top             =   3720
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   3360
      TabIndex        =   20
      Top             =   3720
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   6
      Left            =   3360
      TabIndex        =   16
      Top             =   1800
      Width           =   1000
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   960
      Top             =   6720
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   4680
      TabIndex        =   11
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   4
      Left            =   4680
      TabIndex        =   10
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      Top             =   3720
      Width           =   1000
   End
   Begin VB.ComboBox ComNum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   1200
      List            =   "Form1.frx":0040
      TabIndex        =   6
      Text            =   "COM1"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   1000
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   0
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   1000
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   6720
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3360
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1_RecAllTimeOut 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   6720
   End
   Begin VB.CommandButton Key_OpenCom 
      Caption         =   "打开串口"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Timer Timer_RecTimeOut 
      Enabled         =   0   'False
      Left            =   2760
      Top             =   6720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   460
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   1800
      Width           =   460
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   460
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   3720
      Width           =   460
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   460
      Left            =   5760
      Shape           =   3  'Circle
      Top             =   2760
      Width           =   460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   360
      Left            =   6360
      TabIndex        =   22
      Top             =   3720
      Width           =   3540
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   2880
      X2              =   2880
      Y1              =   1800
      Y2              =   4320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "实时数值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   19
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "智能室内环境监测系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   960
      TabIndex        =   18
      Top             =   240
      Width           =   4815
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "光照"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "当前时间"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   4560
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   360
      Left            =   6360
      TabIndex        =   14
      Top             =   2760
      Width           =   3540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6360
      TabIndex        =   13
      Top             =   1800
      Width           =   3540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "设置上限"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "设置下限"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   7
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "湿度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   900
   End
   Begin VB.Label Label14 
      Caption         =   "串口号"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   5400
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim GucRxDate() As Byte, GucTxDate() As Byte  '将采用动态数组
Dim GucUartDate(30) As Byte
Dim GucSendCount As Integer, Flag_NewDate As Boolean, Flag_UartOpen As Boolean

Dim Flag_SendOther As Boolean
Dim FlagEvent As Integer
Public ADO_Path As String, ADO_Path1 As String
Dim Adocn As ADODB.Connection
Dim FlagNetOPen As Boolean
Dim strkaoqin(100) As String, cntkaoqin As Integer
Dim ka(1) As Byte, kkk As Byte


Private Sub Command1_Click()
    kkk = 1 - kkk
    ka(0) = kkk + 48
    ka(1) = kkk + 48
    MSComm1.Output = ka
End Sub

Private Sub Form_Load()
    Dim mDataBaseName As String
    Dim mCnnStr As String
    Dim cCnn As New ADODB.Connection
    Dim cat As New ADOX.Catalog
    Dim mTable As New ADOX.Table
    Dim mCol As New ADOX.Column
    Dim Adors As New ADODB.Recordset
    Adors.ActiveConnection = Adocn
    Dim str1 As String

    Set Adocn = New ADODB.Connection
'    Dim I As Long, Ls As8 String

    Strold3 = "2F" '初始化变量
            
    mDataBaseName = App.Path & "\数据文件.mdb"
    mCnnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDataBaseName & ";Jet OLEDB:DataBase Password=admin"
    

    
  '  On Error GoTo s1:
        If Dir(mDataBaseName) = "" Then '先查找文件,如果文件不存在,才新建数据库,若存在,则新建数据库会报错
            Call cat.Create(mCnnStr) '创建新带密码的数据库
            '如果数据库已经建,则cat为空,所以必须告诉系统准备使用哪个数据库,其实creat后自动修改activeconnection为最新
            cat.ActiveConnection = mCnnStr

            '新建一个表,名中不能有空格,'-','&'(表是可以,但SQL语句就会报错)
            mTable.Name = "注册表" '创建表,命名为注册表
       
            Set mCol.ParentCatalog = cat
            '创建字段
            With mTable
                
                .Columns.Append "学号", adVarWChar, 100
                .Columns.Append "姓名", adVarWChar, 100
                .Columns.Append "指纹号", adVarWChar, 100
                
                '设置主键
                'Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "日期", "", ""
            End With
            

            
            


        End If
 
        

 
        Set mCol = Nothing
        Set mTable = Nothing
        Set cat = Nothing

        '再新建另外一个表
        mTable.Name = Year(Now) & Month(Now) & Day(Now) '创建表,命名为注册表

        cat.ActiveConnection = mCnnStr
        Set mCol.ParentCatalog = cat
        
        '创建字段
        With mTable
       
            .Columns.Append "温度", adVarWChar, 100
            .Columns.Append "湿度", adVarWChar, 100
    
  .Columns.Append "光照", adVarWChar, 100
           
      
            
            '设置主键
        '    Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "日期", "", ""
        End With
        '生成表
        On Error GoTo SDS
        cat.Tables.Append mTable


      
        

       '打开数据库

SDS:     Adocn.Open mCnnStr
         '读取数据库文件并显示出来
        Adors.ActiveConnection = Adocn
    

       Exit Sub

s1:
       mCol = MsgBox("系统初始化失败,请重启或联系厂商!", vbOKOnly + vbCritical, "警告")
        End
End Sub

Private Sub Info_Change(Index As Integer)
    DelKey.Enabled = False
End Sub

Private Sub Key_OpenCom_Click() '打开串口键按下
    
    
    If Flag_UartOpen = False Then '连接处理

        If ComNum.ListIndex <> -1 Then
            MSComm1.CommPort = ComNum.ListIndex + 1
        Else
            MSComm1.CommPort = 1
        End If
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.InputMode = comInputModeBinary '以文本方式接收
        
        MSComm1.RThreshold = 1 '设置几个字符产生一次oncom事件
        MSComm1.InputLen = 512 '设置ruan冲一次返回的字节数

        MSComm1.InBufferSize = 512 '设置ruan冲区的大小，不能太大和太小，读取时一定会清除接收区，最好和RTHreshold相等
        MSComm1.InBufferCount = 0
        MSComm1.OutBufferSize = 512 '设置ruan冲区的大小，不能太大和太小，读取时一定会清除接收区，最好和RTHreshold相等
        MSComm1.SThreshold = 0
        ''''''''''''''''''''''''''''''''''''其他程序''''''''''''''''''''''''''''''''''''''''''''''''''
        FlagEvent = 0
            
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ErrorCom
        
        MSComm1.PortOpen = True '打开串口,必须放在最前面
        ComNum.Enabled = False
        Key_OpenCom.Caption = "断开串口" '打开串口键的名改为断开串口
        Flag_UartOpen = True
        
        '接收中,每个字节间的最大时间
        Timer_RecTimeOut.Interval = 20 '这个时间可能要通过不同波特率来的

      
        Timer1_RecAllTimeOut.Interval = 3000
        Exit Sub
ErrorCom:
    X = MsgBox("串口不存在或者被占用,请重新连接!", 48, "提示")
    
    Else '断开处理
        MSComm1.PortOpen = False '关闭串口
        '不启动两个超时定时器
        Timer_RecTimeOut.Enabled = False
        Timer1_RecAllTimeOut.Enabled = False
        ComNum.Enabled = True
        Key_OpenCom.Caption = "打开串口"
        Flag_UartOpen = False
        
   
    

        Exit Sub
        
    End If
    
End Sub

Private Sub UartSendErr()
    If FlagEvent = 1 Then
        X = MsgBox("获取参数失败!", 48, "提示")
    ElseIf FlagEvent = 2 Then
        X = MsgBox("保存失败!", 48, "提示")
    ElseIf FlagEvent = 3 Then
        X = MsgBox("删除失败!", 48, "提示")
    ElseIf FlagEvent = 4 Then
        X = MsgBox("增加失败!", 48, "提示")
    End If
    FlagEvent = 0

End Sub

'更新
           ' str1 = "update " & "注册表" & " set 早上签到= " & "'" & "是" & "' " & "where 学号 = " & "'" & "1" & "' " '字符型数据要加单引号
            'Adocn.Execute str1
'测试数据
'fd 01 00 3A 00 00 01 02 3A 00 00 01 03 3A 00 00 01 04 3A 00 00 31 32 33 ff C1 D6 C9 D9 D6 BE ff 34 35 36 ff C0 EE B9 F0 C0 BC ff fe

Public Function Uart_Deal()
    Dim RecCount As Integer, Time1 As Long, Gime(4) As Long, wc1 As String, wc2 As String, wc1flag As String, wc2flag As String
    Dim Adors As New ADODB.Recordset
    Dim SF As Boolean
    Dim str1 As String, str2 As String, str3 As String
    Adors.ActiveConnection = Adocn
    '停止2个超时定时器
    Timer_RecTimeOut.Enabled = False
    Timer1_RecAllTimeOut.Enabled = False
    '得到实际接收到个数
    RecCount = MSComm1.InBufferCount
    Form1.Enabled = True
    '读取数据
    GucRxDate = MSComm1.Input '清空ruan冲区,并清零接收个数

    If RecCount < 4 Then
        Call UartSendErr
        Exit Function
    End If
        
        If GucRxDate(0) = &H53 And GucRxDate(5) = &H45 Then
       str1 = "Insert Into  " + Trim(Str(Year(Now))) + Trim(Str(Month(Now))) + Trim(Str(Day(Now))) + " (温度,湿度,光照)"
       
            Text1(0) = GucRxDate(2)
        
      

            Text1(1) = GucRxDate(3)
       Text1(2) = GucRxDate(4)

          str1 = str1 + "Values('" + Text1(0) + "','" + Text1(1) + "','" + Text1(2) + "')"
          Adocn.Execute str1
          If Text1(0) < Text1(6) Then
          
            Shape1.FillColor = vbBlue
            Label5.ForeColor = vbBlue
            Label5 = "温度过低" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(0) > Text1(4) Then
          
             Shape1.FillColor = vbRed
             Label5.ForeColor = vbRed
             Label5 = "温度过高" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
             
            Else
            Shape1.FillColor = &HC0C0C0
            Label5 = ""
        End If
                    If Text1(1) < Text1(3) Then
                    
            Shape2.FillColor = vbBlue
            Label6.ForeColor = vbBlue
            Label6 = "湿度过低" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(1) > Text1(5) Then
          
            Shape2.FillColor = vbRed
            Label6.ForeColor = vbRed
            Label6 = "湿度过高" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
            Else
            Shape2.FillColor = &HC0C0C0
            Label6 = ""
        End If
         If Text1(2) < Text1(8) Then
         
            Shape3.FillColor = vbBlue
            Label4.ForeColor = vbBlue
            Label4 = "光照度过低" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(2) > Text1(9) Then
          
            Shape3.FillColor = vbRed
            Label4.ForeColor = vbRed
            Label4 = "光照度过高" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
            
            Else
            Shape3.FillColor = &HC0C0C0
            Label4 = ""
        End If
          Exit Function
    End If
         
   



     ''''''''''''''''''''''''''''''''''''数据解析'''''''''''''''''''''''''''''''''''''
        '根据数据及数据个数得到数据是否对与错,然后Flag_NewDate的值告诉你是不是查询命令(是为False)
        '然后再定出错是否要提示对话框(通常,查询命令出错不提示对话框,最好是数据显示?,其他命令提示发送失败与否)】
     

     
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function





Private Sub Timer_RecTimeOut_Timer() '接收字节间超时处理
      Call Uart_Deal
End Sub
Private Sub Timer1_RecAllTimeOut_Timer() '接收总处理
      Call Uart_Deal
End Sub


Private Sub UartSend() '发送
On Error GoTo gk
    '重新定义动态数组
    ReDim GucTxDate(0 To GucSendCount - 1) As Byte
    '数据复制
    For i = 0 To GucSendCount - 1
        GucTxDate(i) = GucUartDate(i)
    Next i
    '发送前先读取缓冲区并清零
    GucRxDate = MSComm1.Input '清空ruan冲区,并清零接收个数
    '启动串口发送
    MSComm1.Output = GucTxDate

    '启动总的超时定时器
    Timer1_RecAllTimeOut.Enabled = True
    Exit Sub
gk:
    Call Key_OpenCom_Click
End Sub



Private Sub ChangeKey_Click() '修改考勤时间
    Dim Adors As New ADODB.Recordset, str1 As String
    Adors.ActiveConnection = Adocn
    '"delete from 数据表" (将数据表所有记录删除)
    str1 = "delete from " + "时间"
    Adocn.Execute str1
    str1 = "Insert Into " + "时间" + "(早上起始时间,早上结束时间,中午起始时间,中午结束时间) "
    str1 = str1 + "Values('" + Time(0) + "','" + Time(1) + "','" + Time(2) + "','" + Time(3) + "')"
     Adocn.Execute str1
    X = MsgBox("保存成功!", 48, "提示")
End Sub


Private Sub AddKey_Click() '增加按键

    GucUartDate(0) = &H58
    GucUartDate(1) = &H4

    GucSendCount = 2
    Timer1_RecAllTimeOut.Interval = 60000 '接收超时
    Form1.Enabled = False
    FlagEvent = 4
    Call UartSend
    
    
End Sub



'串口接收中断(每收一个字节中断一次）
Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent = 2 And MSComm1.InBufferCount <> 0 Then '如果是接收到RThreshol个字符而产生该事件的话
        '重新置位接收超时,试过,这三条语句不会影响到,即下面发多少个字节,这串口接收中断就几次,也就复位几次接收超时定时器
        Timer_RecTimeOut.Enabled = False
        Timer_RecTimeOut.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If Hour(Now) = 0 And Minute(Now) = 0 Then
    
    End If
End Sub

Private Sub Timer2_Timer()
    Label7 = "当前时间:" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Sub
