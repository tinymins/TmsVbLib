Attribute VB_Name = "modXP"

'�����뻶ӭ����ת���������໥̽��,���뱣��������˵�� 2005/12/31
'����:�γ��� �Ͼ�  ������ҳhttp://www.asanscape.com
'����BLOG: http://blog.csdn.net/asanscape     VBȺ:12960265  Email:asangray@163.com

'�������е�xp.res����ֱ�Ӽ����������̽��б�����ʹ��߱�XP���
'ע�������Դ�ļ�ʱҪͬʱ���뱾ģ��,  ���ù��̴�Sub Main()����,�����޳�ʼ������
'InitCommonControls����������comctl32.dll(�汾5)��,������ʹ��,��Ҫʹ��InitCommonControlsEx

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Sub Main()
	InitCommonControlsVB
    Form1.Show
End Sub



