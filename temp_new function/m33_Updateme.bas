'VERSION 1.0 CLASS
'Begin
'  MultiUse = -1  'True
'End
'Attribute VB_Name = "CatiaUpdateForm"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = False
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'Public CATIA As Application
'Public chkTM
'Public txtTM As MSForms.TextBox
'Public chkPath As MSForms.CheckBox
'Public cmdOK As MSForms.CommandButton
'Public cmdCancel As MSForms.CommandButton

Private Const itemgap = 10
Private Const fmargin = 10
Private Const bmargin = 10

Public chkTM, chkPath As MSForms.CheckBox
Public txtTM As MSForms.TextBox
Public cmdCancel, cmdok As MSForms.CommandButton
Public date_path
Public dateFMT

 

Private Sub UserForm_Initialize()
    ' ���ô����������
    Me.Caption = "����Ÿ��º�·��ѡ��"
    Me.Width = 200
    Me.Height = 300
    Me.StartUpPosition = 1
   
    Set chkTM = Me.Controls.Add("Forms.CheckBox.1", "cjk", True)
    With chkTM
        .Name = "chkTM"
        .Caption = "�Ƿ����CATIA�����ʱ���"
        .Left = fmargin
        .Top = fmargin
        .Width = 180
        .Height = 20
    End With

    
    Set txtTM = Me.Controls.Add("Forms.TextBox.1", "cjk", True)
    With txtTM
        .Name = "txtTM"
        .Text = "���������ڸ�ʽ�� ��=d��day��ʱ=hour��h����=min��i��Ĭ��Ϊ��"
        .Enabled = False
        .Left = 30
        .Top = 90
        .Width = 200
        .Height = 20
    End With
    
    Set chkPath = Me.Controls.Add("Forms.CheckBox.1")
    With chkPath
        .Name = "chkPath"
        .Caption = "��������ǰ·��"
        .Left = 30
        .Top = 130
        .Width = 150
        .Height = 20
    End With
    
    Set cmdok = Me.Controls.Add("Forms.CommandButton.1")
    With cmdok
        .Name = "cmdOK"
        .Caption = "ȷ��"
        .Left = 50
        .Top = 200
        .Width = 80
        .Height = 30
    End With
    
    Set cmdCancel = Me.Controls.Add("Forms.CommandButton.1")
    With cmdCancel
        .Name = "cmdCancel"
        .Caption = "ȡ��"
        .Left = 180
        .Top = 200
        .Width = 80
        .Height = 30
    End With
End Sub

' �¼�����������ʹ����ȷ������Լ��
Private Sub chkTM_Click()
    txtTM.Enabled = chkTM.Value
End Sub

Private Sub cmdOK_Click()
 Debug.Print "i���ǵö�"
    Dim UpdateTimestamp As Boolean
    Dim ExportToCurrentPath As Boolean
    UpdateTimestamp = chkTM.Value
    ExportToCurrentPath = chkPath.Value
    Set date_path = Array(UpdateTimestamp, Trim(txtTM.Text), ExportToCurrentPath)
    Debug.Print date_path(0)
    MsgBox date_path(0)
    MsgBox date_path(1)
    MsgBox date_path(2)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' ʵ�ʵ�CATIA��������
Private Sub ProcessUpdates(UpdateTimestamp As Boolean, ExportToCurrentPath As Boolean, DateFormat As String)
    
    MsgBox "��������: " & Err.Description, vbCritical
End Sub
