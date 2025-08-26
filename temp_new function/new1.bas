VERSION 5.00
Begin VB.Form DatePathForm 
    Caption         =   "���ں�·������"
    ClientHeight    =   2100
    ClientLeft      =   45
    ClientTop       =   330
    ClientWidth     =   4680
    Icon            =   "DatePathForm.frx":0000
    LinkTopic       =   "Form1"
    ScaleHeight     =   2100
    ScaleWidth      =   4680
    StartUpPosition =   1  '����������
    Begin VB.CommandButton cmdExecute 
        Caption         =   "ִ�в���"
        Height          =   495
        Left            =   1440
        TabIndex        =   3
        Top             =   1440
        Width           =   1695
    End
    Begin VB.CheckBox chkPath 
        Caption         =   "����·��"
        Height          =   375
        Left            =   720
        TabIndex        =   2
        Top             =   960
        Width           =   2895
    End
    Begin VB.TextBox txtDate 
        Enabled         =   0   'False
        Height          =   375
        Left            =   1800
        TabIndex        =   1
        Text            =   "YYYY-MM-DD"
        Top             =   480
        Width           =   2175
    End
    Begin VB.CheckBox chkDate 
        Caption         =   "��������"
        Height          =   375
        Left            =   720
        TabIndex        =   0
        Top             =   480
        Width           =   1095
    End
    Begin VB.Label Label1 
        Caption         =   "��ѡ����Ҫִ�еĲ�����"
        Height          =   375
        Left            =   720
        TabIndex        =   4
        Top             =   120
        Width           =   2895
    End
End
Attribute VB_Name = "DatePathForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ִ�а�ť����¼�
Private Sub cmdExecute_Click()
    Dim updateDate As Boolean
    Dim updatePath As Boolean
    Dim dateValue As String
    
    ' ��ȡ��ѡ��״̬
    updateDate = chkDate.Value
    updatePath = chkPath.Value
    
    ' ��֤����ѡ��
    If updateDate Then
        ' ��������ı����Ƿ���ֵ
        If Trim(txtDate.Text) = "" Or txtDate.Text = "YYYY-MM-DD" Then
            MsgBox "���������ı�����������Ч������ֵ", vbExclamation, "�������"
            txtDate.SetFocus
            Exit Sub
        End If
        
        ' �򵥵����ڸ�ʽ��֤
        If Not IsDate(txtDate.Text) Then
            MsgBox "��������Ч�����ڸ�ʽ���磺2023-12-31��", vbExclamation, "��ʽ����"
            txtDate.SetFocus
            Exit Sub
        End If
        
        ' ��������ֵ
        dateValue = txtDate.Text
    End If
    
    ' ִ�в���
    ExecuteOperations updateDate, updatePath, dateValue
End Sub

' ִ�о�������Ĺ���
Private Sub ExecuteOperations(needUpdateDate As Boolean, needUpdatePath As Boolean, newDate As String)
    Dim resultMsg As String
    resultMsg = "��������ɣ�" & vbCrLf & vbCrLf
    
    ' �������ڸ���
    If needUpdateDate Then
        resultMsg = resultMsg & "? �Ѹ�������Ϊ: " & Format(newDate, "yyyy��mm��dd��") & vbCrLf
        ' ���������ʵ�ʸ������ڵĴ���
        ' ����: UpdateSystemDate newDate
    Else
        resultMsg = resultMsg & "? δ��������" & vbCrLf
    End If
    
    ' ����·������
    If needUpdatePath Then
        resultMsg = resultMsg & "? �Ѹ���·��" & vbCrLf
        ' ���������ʵ�ʸ���·���Ĵ���
        ' ����: UpdateFilePath
    Else
        resultMsg = resultMsg & "? δ����·��" & vbCrLf
    End If
    
    ' ��ʾ�������
    MsgBox resultMsg, vbInformation, "�������"
End Sub

' ���ڸ�ѡ��״̬�仯�¼�
Private Sub chkDate_Click()
    ' ����ѡ���ڸ�ѡ��ʱ�����ı��򣬷������
    txtDate.Enabled = chkDate.Value
    
    ' ����ǹ�ѡ״̬���ı�����Ĭ����ʾ���֣������
    If chkDate.Value And txtDate.Text = "YYYY-MM-DD" Then
        txtDate.Text = ""
    ElseIf Not chkDate.Value And txtDate.Text = "" Then
        ' ���ȡ����ѡ���ı���Ϊ�գ���ָ���ʾ����
        txtDate.Text = "YYYY-MM-DD"
    End If
End Sub

' �����ʼ���¼�
Private Sub UserForm_Initialize()
    ' ���ó�ʼ״̬
    txtDate.Enabled = False
    chkDate.Value = False
    chkPath.Value = False
End Sub
