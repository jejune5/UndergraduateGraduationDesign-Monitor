VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "�������ڻ������ϵͳ"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   FillColor       =   &H00C0C0FF&
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "�򿪴���"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ʵʱ��ֵ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�������ڻ������ϵͳ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǰʱ��"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ʪ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ں�"
      BeginProperty Font 
         Name            =   "����"
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


Dim GucRxDate() As Byte, GucTxDate() As Byte  '�����ö�̬����
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

    Strold3 = "2F" '��ʼ������
            
    mDataBaseName = App.Path & "\�����ļ�.mdb"
    mCnnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mDataBaseName & ";Jet OLEDB:DataBase Password=admin"
    

    
  '  On Error GoTo s1:
        If Dir(mDataBaseName) = "" Then '�Ȳ����ļ�,����ļ�������,���½����ݿ�,������,���½����ݿ�ᱨ��
            Call cat.Create(mCnnStr) '�����´���������ݿ�
            '������ݿ��Ѿ���,��catΪ��,���Ա������ϵͳ׼��ʹ���ĸ����ݿ�,��ʵcreat���Զ��޸�activeconnectionΪ����
            cat.ActiveConnection = mCnnStr

            '�½�һ����,���в����пո�,'-','&'(���ǿ���,��SQL���ͻᱨ��)
            mTable.Name = "ע���" '������,����Ϊע���
       
            Set mCol.ParentCatalog = cat
            '�����ֶ�
            With mTable
                
                .Columns.Append "ѧ��", adVarWChar, 100
                .Columns.Append "����", adVarWChar, 100
                .Columns.Append "ָ�ƺ�", adVarWChar, 100
                
                '��������
                'Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "����", "", ""
            End With
            

            
            


        End If
 
        

 
        Set mCol = Nothing
        Set mTable = Nothing
        Set cat = Nothing

        '���½�����һ����
        mTable.Name = Year(Now) & Month(Now) & Day(Now) '������,����Ϊע���

        cat.ActiveConnection = mCnnStr
        Set mCol.ParentCatalog = cat
        
        '�����ֶ�
        With mTable
       
            .Columns.Append "�¶�", adVarWChar, 100
            .Columns.Append "ʪ��", adVarWChar, 100
    
  .Columns.Append "����", adVarWChar, 100
           
      
            
            '��������
        '    Keys.Append "PrimaryKey", ADOX.KeyTypeEnum.adKeyPrimary, "����", "", ""
        End With
        '���ɱ�
        On Error GoTo SDS
        cat.Tables.Append mTable


      
        

       '�����ݿ�

SDS:     Adocn.Open mCnnStr
         '��ȡ���ݿ��ļ�����ʾ����
        Adors.ActiveConnection = Adocn
    

       Exit Sub

s1:
       mCol = MsgBox("ϵͳ��ʼ��ʧ��,����������ϵ����!", vbOKOnly + vbCritical, "����")
        End
End Sub

Private Sub Info_Change(Index As Integer)
    DelKey.Enabled = False
End Sub

Private Sub Key_OpenCom_Click() '�򿪴��ڼ�����
    
    
    If Flag_UartOpen = False Then '���Ӵ���

        If ComNum.ListIndex <> -1 Then
            MSComm1.CommPort = ComNum.ListIndex + 1
        Else
            MSComm1.CommPort = 1
        End If
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.InputMode = comInputModeBinary '���ı���ʽ����
        
        MSComm1.RThreshold = 1 '���ü����ַ�����һ��oncom�¼�
        MSComm1.InputLen = 512 '����ruan��һ�η��ص��ֽ���

        MSComm1.InBufferSize = 512 '����ruan�����Ĵ�С������̫���̫С����ȡʱһ�����������������ú�RTHreshold���
        MSComm1.InBufferCount = 0
        MSComm1.OutBufferSize = 512 '����ruan�����Ĵ�С������̫���̫С����ȡʱһ�����������������ú�RTHreshold���
        MSComm1.SThreshold = 0
        ''''''''''''''''''''''''''''''''''''��������''''''''''''''''''''''''''''''''''''''''''''''''''
        FlagEvent = 0
            
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error GoTo ErrorCom
        
        MSComm1.PortOpen = True '�򿪴���,���������ǰ��
        ComNum.Enabled = False
        Key_OpenCom.Caption = "�Ͽ�����" '�򿪴��ڼ�������Ϊ�Ͽ�����
        Flag_UartOpen = True
        
        '������,ÿ���ֽڼ�����ʱ��
        Timer_RecTimeOut.Interval = 20 '���ʱ�����Ҫͨ����ͬ����������

      
        Timer1_RecAllTimeOut.Interval = 3000
        Exit Sub
ErrorCom:
    X = MsgBox("���ڲ����ڻ��߱�ռ��,����������!", 48, "��ʾ")
    
    Else '�Ͽ�����
        MSComm1.PortOpen = False '�رմ���
        '������������ʱ��ʱ��
        Timer_RecTimeOut.Enabled = False
        Timer1_RecAllTimeOut.Enabled = False
        ComNum.Enabled = True
        Key_OpenCom.Caption = "�򿪴���"
        Flag_UartOpen = False
        
   
    

        Exit Sub
        
    End If
    
End Sub

Private Sub UartSendErr()
    If FlagEvent = 1 Then
        X = MsgBox("��ȡ����ʧ��!", 48, "��ʾ")
    ElseIf FlagEvent = 2 Then
        X = MsgBox("����ʧ��!", 48, "��ʾ")
    ElseIf FlagEvent = 3 Then
        X = MsgBox("ɾ��ʧ��!", 48, "��ʾ")
    ElseIf FlagEvent = 4 Then
        X = MsgBox("����ʧ��!", 48, "��ʾ")
    End If
    FlagEvent = 0

End Sub

'����
           ' str1 = "update " & "ע���" & " set ����ǩ��= " & "'" & "��" & "' " & "where ѧ�� = " & "'" & "1" & "' " '�ַ�������Ҫ�ӵ�����
            'Adocn.Execute str1
'��������
'fd 01 00 3A 00 00 01 02 3A 00 00 01 03 3A 00 00 01 04 3A 00 00 31 32 33 ff C1 D6 C9 D9 D6 BE ff 34 35 36 ff C0 EE B9 F0 C0 BC ff fe

Public Function Uart_Deal()
    Dim RecCount As Integer, Time1 As Long, Gime(4) As Long, wc1 As String, wc2 As String, wc1flag As String, wc2flag As String
    Dim Adors As New ADODB.Recordset
    Dim SF As Boolean
    Dim str1 As String, str2 As String, str3 As String
    Adors.ActiveConnection = Adocn
    'ֹͣ2����ʱ��ʱ��
    Timer_RecTimeOut.Enabled = False
    Timer1_RecAllTimeOut.Enabled = False
    '�õ�ʵ�ʽ��յ�����
    RecCount = MSComm1.InBufferCount
    Form1.Enabled = True
    '��ȡ����
    GucRxDate = MSComm1.Input '���ruan����,��������ո���

    If RecCount < 4 Then
        Call UartSendErr
        Exit Function
    End If
        
        If GucRxDate(0) = &H53 And GucRxDate(5) = &H45 Then
       str1 = "Insert Into  " + Trim(Str(Year(Now))) + Trim(Str(Month(Now))) + Trim(Str(Day(Now))) + " (�¶�,ʪ��,����)"
       
            Text1(0) = GucRxDate(2)
        
      

            Text1(1) = GucRxDate(3)
       Text1(2) = GucRxDate(4)

          str1 = str1 + "Values('" + Text1(0) + "','" + Text1(1) + "','" + Text1(2) + "')"
          Adocn.Execute str1
          If Text1(0) < Text1(6) Then
          
            Shape1.FillColor = vbBlue
            Label5.ForeColor = vbBlue
            Label5 = "�¶ȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(0) > Text1(4) Then
          
             Shape1.FillColor = vbRed
             Label5.ForeColor = vbRed
             Label5 = "�¶ȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
             
            Else
            Shape1.FillColor = &HC0C0C0
            Label5 = ""
        End If
                    If Text1(1) < Text1(3) Then
                    
            Shape2.FillColor = vbBlue
            Label6.ForeColor = vbBlue
            Label6 = "ʪ�ȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(1) > Text1(5) Then
          
            Shape2.FillColor = vbRed
            Label6.ForeColor = vbRed
            Label6 = "ʪ�ȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
            Else
            Shape2.FillColor = &HC0C0C0
            Label6 = ""
        End If
         If Text1(2) < Text1(8) Then
         
            Shape3.FillColor = vbBlue
            Label4.ForeColor = vbBlue
            Label4 = "���նȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
          ElseIf Text1(2) > Text1(9) Then
          
            Shape3.FillColor = vbRed
            Label4.ForeColor = vbRed
            Label4 = "���նȹ���" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
            
            
            Else
            Shape3.FillColor = &HC0C0C0
            Label4 = ""
        End If
          Exit Function
    End If
         
   



     ''''''''''''''''''''''''''''''''''''���ݽ���'''''''''''''''''''''''''''''''''''''
        '�������ݼ����ݸ����õ������Ƿ�����,Ȼ��Flag_NewDate��ֵ�������ǲ��ǲ�ѯ����(��ΪFalse)
        'Ȼ���ٶ������Ƿ�Ҫ��ʾ�Ի���(ͨ��,��ѯ���������ʾ�Ի���,�����������ʾ?,����������ʾ����ʧ�����)��
     

     
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Function





Private Sub Timer_RecTimeOut_Timer() '�����ֽڼ䳬ʱ����
      Call Uart_Deal
End Sub
Private Sub Timer1_RecAllTimeOut_Timer() '�����ܴ���
      Call Uart_Deal
End Sub


Private Sub UartSend() '����
On Error GoTo gk
    '���¶��嶯̬����
    ReDim GucTxDate(0 To GucSendCount - 1) As Byte
    '���ݸ���
    For i = 0 To GucSendCount - 1
        GucTxDate(i) = GucUartDate(i)
    Next i
    '����ǰ�ȶ�ȡ������������
    GucRxDate = MSComm1.Input '���ruan����,��������ո���
    '�������ڷ���
    MSComm1.Output = GucTxDate

    '�����ܵĳ�ʱ��ʱ��
    Timer1_RecAllTimeOut.Enabled = True
    Exit Sub
gk:
    Call Key_OpenCom_Click
End Sub



Private Sub ChangeKey_Click() '�޸Ŀ���ʱ��
    Dim Adors As New ADODB.Recordset, str1 As String
    Adors.ActiveConnection = Adocn
    '"delete from ���ݱ�" (�����ݱ����м�¼ɾ��)
    str1 = "delete from " + "ʱ��"
    Adocn.Execute str1
    str1 = "Insert Into " + "ʱ��" + "(������ʼʱ��,���Ͻ���ʱ��,������ʼʱ��,�������ʱ��) "
    str1 = str1 + "Values('" + Time(0) + "','" + Time(1) + "','" + Time(2) + "','" + Time(3) + "')"
     Adocn.Execute str1
    X = MsgBox("����ɹ�!", 48, "��ʾ")
End Sub


Private Sub AddKey_Click() '���Ӱ���

    GucUartDate(0) = &H58
    GucUartDate(1) = &H4

    GucSendCount = 2
    Timer1_RecAllTimeOut.Interval = 60000 '���ճ�ʱ
    Form1.Enabled = False
    FlagEvent = 4
    Call UartSend
    
    
End Sub



'���ڽ����ж�(ÿ��һ���ֽ��ж�һ�Σ�
Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent = 2 And MSComm1.InBufferCount <> 0 Then '����ǽ��յ�RThreshol���ַ����������¼��Ļ�
        '������λ���ճ�ʱ,�Թ�,��������䲻��Ӱ�쵽,�����淢���ٸ��ֽ�,�⴮�ڽ����жϾͼ���,Ҳ�͸�λ���ν��ճ�ʱ��ʱ��
        Timer_RecTimeOut.Enabled = False
        Timer_RecTimeOut.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
    If Hour(Now) = 0 And Minute(Now) = 0 Then
    
    End If
End Sub

Private Sub Timer2_Timer()
    Label7 = "��ǰʱ��:" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now)
End Sub
