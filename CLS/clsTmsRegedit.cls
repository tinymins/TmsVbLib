VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTmsRegedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FILETIME ' 8 Bytes
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
'���ڲ���ע�����API�������� -begin-
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long                                                                                 '   Note   that   if   you   declare   the   lpData   parameter   as   String,   you   must   pass   it   By   Value.
Private Declare Function RegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpValue As String, lpcbValue As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
'���ڲ���ע����ĳ������� -begin-
'Ԥ����ĸ����
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_DYN_DATA = &H80000006
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_USERS = &H80000003
'Ԥ����ļ�ֵ����
Private Const REG_NONE = 0
Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_LITTLE_ENDIAN = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_LINK = 6
Private Const REG_MULTI_SZ = 7
Private Const REG_RESOURCE_LIST = 8
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10
'#If WinNT Then
'ACCESS_MASK
Private Const DELETE = (&H10000)
Private Const READ_CONTROL = (&H20000)
Private Const WRITE_DAC = (&H40000)
Private Const WRITE_OWNER = (&H80000)
Private Const SYNCHRONIZE = (&H100000)
Private Const STANDARD_RIGHTS_REQUIRED = (&HF0000)
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL = (&H1F0000)
Private Const SPECIFIC_RIGHTS_ALL = (&HFFFF)
'��ע���ļ��ķ�ʽ
Private Const KEY_EVENT = &H1
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_EXECUTE = (KEY_READ)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'#End If
'ע���ļ���������
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_REGISTRY_RECOVERED = 1014&
Private Const ERROR_REGISTRY_CORRUPT = 1015&
Private Const ERROR_REGISTRY_IO_FAILED = 1016&
Private Const ERROR_NOT_REGISTRY_FILE = 1017&
Private Const ERROR_KEY_DELETED = 1018&
Private Const ERROR_NO_LOG_SPACE = 1019&
Private Const ERROR_KEY_HAS_CHILDREN = 1020&
Private Const ERROR_CHILD_MUST_BE_VOLATILE = 1021&
Private Const ERROR_RXACT_INVALID_STATE = 1369&
Private Const ERROR_NO_MORE_ITEMS = 259&
' �Զ���ע���ļ�����
Private Const REGAGENT_NOKEY = -1002
Private Const REGAGENT_NOSUBKEY = -1003
'���ڲ���ע����ĳ������� -end-

Private glStatus As Long

'���ڲ���ע����ĳ������� -begin-
'Ԥ����ĸ����
Public Property Get CONST_HKEY_CLASSES_ROOT()
    CONST_HKEY_CLASSES_ROOT = HKEY_CLASSES_ROOT
End Property
Public Property Get CONST_HKEY_CURRENT_CONFIG()
    CONST_HKEY_CURRENT_CONFIG = HKEY_CURRENT_CONFIG
End Property
Public Property Get CONST_HKEY_CURRENT_USER()
    CONST_HKEY_CURRENT_USER = HKEY_CURRENT_USER
End Property
Public Property Get CONST_HKEY_DYN_DATA()
    CONST_HKEY_DYN_DATA = HKEY_DYN_DATA
End Property
Public Property Get CONST_HKEY_LOCAL_MACHINE()
    CONST_HKEY_LOCAL_MACHINE = HKEY_LOCAL_MACHINE
End Property
Public Property Get CONST_HKEY_PERFORMANCE_DATA()
    CONST_HKEY_PERFORMANCE_DATA = HKEY_PERFORMANCE_DATA
End Property
Public Property Get CONST_HKEY_USERS()
    CONST_HKEY_USERS = HKEY_USERS
End Property
'Ԥ����ļ�ֵ����
Public Property Get CONST_REG_NONE()
    CONST_REG_NONE = REG_NONE
End Property
Public Property Get CONST_REG_SZ()
    CONST_REG_SZ = REG_SZ
End Property
Public Property Get CONST_REG_EXPAND_SZ()
    CONST_REG_EXPAND_SZ = REG_EXPAND_SZ
End Property
Public Property Get CONST_REG_BINARY()
    CONST_REG_BINARY = REG_BINARY
End Property
Public Property Get CONST_REG_DWORD()
    CONST_REG_DWORD = REG_DWORD
End Property
Public Property Get CONST_REG_DWORD_LITTLE_ENDIAN()
    CONST_REG_DWORD_LITTLE_ENDIAN = REG_DWORD_LITTLE_ENDIAN
End Property
Public Property Get CONST_REG_DWORD_BIG_ENDIAN()
    CONST_REG_DWORD_BIG_ENDIAN = REG_DWORD_BIG_ENDIAN
End Property
Public Property Get CONST_REG_LINK()
    CONST_REG_LINK = REG_LINK
End Property
Public Property Get CONST_REG_MULTI_SZ()
    CONST_REG_MULTI_SZ = REG_MULTI_SZ
End Property
Public Property Get CONST_REG_RESOURCE_LIST()
    CONST_REG_RESOURCE_LIST = REG_RESOURCE_LIST
End Property
Public Property Get CONST_REG_FULL_RESOURCE_DESCRIPTOR()
    CONST_REG_FULL_RESOURCE_DESCRIPTOR = REG_FULL_RESOURCE_DESCRIPTOR
End Property
Public Property Get CONST_REG_RESOURCE_REQUIREMENTS_LIST()
    CONST_REG_RESOURCE_REQUIREMENTS_LIST = REG_RESOURCE_REQUIREMENTS_LIST
End Property
'#If WinNT Then
'ACCESS MASK
Public Property Get CONST_DELETE()
    CONST_DELETE = DELETE
End Property
Public Property Get CONST_READ_CONTROL()
    CONST_READ_CONTROL = READ_CONTROL
End Property
Public Property Get CONST_WRITE_DAC()
    CONST_WRITE_DAC = WRITE_DAC
End Property
Public Property Get CONST_WRITE_OWNER()
    CONST_WRITE_OWNER = WRITE_OWNER
End Property
Public Property Get CONST_SYNCHRONIZE()
    CONST_SYNCHRONIZE = SYNCHRONIZE
End Property
Public Property Get CONST_STANDARD_RIGHTS_REQUIRED()
    CONST_STANDARD_RIGHTS_REQUIRED = STANDARD_RIGHTS_REQUIRED
End Property
Public Property Get CONST_STANDARD_RIGHTS_READ()
    CONST_STANDARD_RIGHTS_READ = STANDARD_RIGHTS_READ
End Property
Public Property Get CONST_STANDARD_RIGHTS_WRITE()
    CONST_STANDARD_RIGHTS_WRITE = STANDARD_RIGHTS_WRITE
End Property
Public Property Get CONST_STANDARD_RIGHTS_EXECUTE()
    CONST_STANDARD_RIGHTS_EXECUTE = STANDARD_RIGHTS_EXECUTE
End Property
Public Property Get CONST_STANDARD_RIGHTS_ALL()
    CONST_STANDARD_RIGHTS_ALL = STANDARD_RIGHTS_ALL
End Property
Public Property Get CONST_SPECIFIC_RIGHTS_ALL()
    CONST_SPECIFIC_RIGHTS_ALL = SPECIFIC_RIGHTS_ALL
End Property
'��ע���ļ��ķ�ʽ
Public Property Get CONST_KEY_EVENT()
    CONST_KEY_EVENT = KEY_EVENT
End Property
Public Property Get CONST_KEY_NOTIFY()
    CONST_KEY_NOTIFY = KEY_NOTIFY
End Property
Public Property Get CONST_KEY_QUERY_VALUE()
    CONST_KEY_QUERY_VALUE = KEY_QUERY_VALUE
End Property
Public Property Get CONST_KEY_SET_VALUE()
    CONST_KEY_SET_VALUE = KEY_SET_VALUE
End Property
Public Property Get CONST_KEY_CREATE_SUB_KEY()
    CONST_KEY_CREATE_SUB_KEY = KEY_CREATE_SUB_KEY
End Property
Public Property Get CONST_KEY_ENUMERATE_SUB_KEYS()
    CONST_KEY_ENUMERATE_SUB_KEYS = KEY_ENUMERATE_SUB_KEYS
End Property
Public Property Get CONST_KEY_CREATE_LINK()
    CONST_KEY_CREATE_LINK = KEY_CREATE_LINK
End Property
Public Property Get CONST_KEY_READ()
    CONST_KEY_READ = KEY_READ
End Property
Public Property Get CONST_KEY_EXECUTE()
    CONST_KEY_EXECUTE = KEY_EXECUTE
End Property
Public Property Get CONST_KEY_WRITE()
    CONST_KEY_WRITE = KEY_WRITE
End Property
Public Property Get CONST_KEY_ALL_ACCESS()
    CONST_KEY_ALL_ACCESS = KEY_ALL_ACCESS
End Property
'#End If
'ע���ļ���������
Public Property Get CONST_ERROR_SUCCESS()
    CONST_ERROR_SUCCESS = ERROR_SUCCESS
End Property
Public Property Get CONST_ERROR_BADDB()
    CONST_ERROR_BADDB = ERROR_BADDB
End Property
Public Property Get CONST_ERROR_BADKEY()
    CONST_ERROR_BADKEY = ERROR_BADKEY
End Property
Public Property Get CONST_ERROR_CANTOPEN()
    CONST_ERROR_CANTOPEN = ERROR_CANTOPEN
End Property
Public Property Get CONST_ERROR_CANTREAD()
    CONST_ERROR_CANTREAD = ERROR_CANTREAD
End Property
Public Property Get CONST_ERROR_CANTWRITE()
    CONST_ERROR_CANTWRITE = ERROR_CANTWRITE
End Property
Public Property Get CONST_ERROR_REGISTRY_RECOVERED()
    CONST_ERROR_REGISTRY_RECOVERED = ERROR_REGISTRY_RECOVERED
End Property
Public Property Get CONST_ERROR_REGISTRY_CORRUPT()
    CONST_ERROR_REGISTRY_CORRUPT = ERROR_REGISTRY_CORRUPT
End Property
Public Property Get CONST_ERROR_REGISTRY_IO_FAILED()
    CONST_ERROR_REGISTRY_IO_FAILED = ERROR_REGISTRY_IO_FAILED
End Property
Public Property Get CONST_ERROR_NOT_REGISTRY_FILE()
    CONST_ERROR_NOT_REGISTRY_FILE = ERROR_NOT_REGISTRY_FILE
End Property
Public Property Get CONST_ERROR_KEY_DELETED()
    CONST_ERROR_KEY_DELETED = ERROR_KEY_DELETED
End Property
Public Property Get CONST_ERROR_NO_LOG_SPACE()
    CONST_ERROR_NO_LOG_SPACE = ERROR_NO_LOG_SPACE
End Property
Public Property Get CONST_ERROR_KEY_HAS_CHILDREN()
    CONST_ERROR_KEY_HAS_CHILDREN = ERROR_KEY_HAS_CHILDREN
End Property
Public Property Get CONST_ERROR_CHILD_MUST_BE_VOLATILE()
    CONST_ERROR_CHILD_MUST_BE_VOLATILE = ERROR_CHILD_MUST_BE_VOLATILE
End Property
Public Property Get CONST_ERROR_RXACT_INVALID_STATE()
    CONST_ERROR_RXACT_INVALID_STATE = ERROR_RXACT_INVALID_STATE
End Property
Public Property Get CONST_ERROR_NO_MORE_ITEMS()
    CONST_ERROR_NO_MORE_ITEMS = ERROR_NO_MORE_ITEMS
End Property
' �Զ���ע���ļ�����
Public Property Get CONST_REGAGENT_NOKEY()
    CONST_REGAGENT_NOKEY = REGAGENT_NOKEY
End Property
Public Property Get CONST_REGAGENT_NOSUBKEY()
    CONST_REGAGENT_NOSUBKEY = REGAGENT_NOSUBKEY
End Property
'���ڲ���ע����ĳ������� -end-

Public Property Get Status()
    Status = glStatus
End Property

Public Function gfsGetKeyStringValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As String
    '�� �ܣ���ע�����ȡ�ô�ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' ���: gfsGetKeyStringValue String ȡ�õ�ע�����ֵ
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    Dim llBufferSize As Long '���ȡ���Ĵ�ֵ����
    Dim lsKeyValue As String '��Ŷ�ȡ�Ĵ�ֵ
    
    'Ԥ����Ϊ��
    gfsGetKeyStringValue = Empty
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
        
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then '�ɹ���ȡ���ȡ�ִ��Ĵ�ֵ��С
        glStatus = RegQueryValueEx(llKeyID, psSubKey, 0&, REG_SZ, 0&, llBufferSize)
        If llBufferSize < 2 Then '��ֵ
            glStatus = RegCloseKey(llKeyID)
        Else '��ֵ����ʽ��ȡ��ֵ
            lsKeyValue = String(llBufferSize + 1, " ")
            glStatus = RegQueryValueEx(llKeyID, psSubKey, 0&, REG_SZ, ByVal lsKeyValue, llBufferSize)
            If glStatus = ERROR_SUCCESS Then
                gfsGetKeyStringValue = Left$(lsKeyValue, llBufferSize - 1)
            End If
            glStatus = RegCloseKey(llKeyID)
        End If
    End If
End Function

Public Function gflGetKeyBinaryValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As Long
    '�� �ܣ���ע�����ȡ�ö�����ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' ���: gflGetKeyBinaryValue Long ȡ�õ�ע���������ֵ
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    Dim llKeyValue As Long '��Ŷ�ȡ�Ķ�����ֵ
    
    'Ԥ����Ϊ��
    gflGetKeyBinaryValue = Empty
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
    
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then '�ɹ���ȡ������ֵ
        glStatus = RegQueryValueEx(llKeyID, psSubKey, 0&, REG_BINARY, llKeyValue, Len(llKeyValue))
        If glStatus = ERROR_SUCCESS Then
            gflGetKeyBinaryValue = llKeyValue
        End If
        glStatus = RegCloseKey(llKeyID)
    End If
End Function
Public Function gflGetKeyDwordValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As Long
    'ok
    '�� �ܣ���ע�����ȡ��˫��ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' ���: gflGetKeyDwordValue Long ȡ�õ�ע���˫��ֵ
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    Dim llKeyValue As Long '��Ŷ�ȡ��˫��ֵ
    
    'Ԥ����Ϊ��
    gflGetKeyDwordValue = Empty
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
    
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then '�ɹ���ȡ˫��ֵ
        glStatus = RegQueryValueEx(llKeyID, psSubKey, 0&, REG_DWORD, llKeyValue, Len(llKeyValue))
        If glStatus = ERROR_SUCCESS Then
            gflGetKeyDwordValue = llKeyValue
        End If
        glStatus = RegCloseKey(llKeyID)
    End If
End Function
Public Sub gpvSetKeyStringValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal psKeyValue As String)
    'ok
    '�� �ܣ�����ע����еĴ�ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' psKeyValue String Ҫ���õĴ�ֵ
    ' ���: ��
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Sub
    End If
    
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    If glStatus = ERROR_SUCCESS Then '�ɹ�������ֵ
        If Len(psKeyValue) = 0 Then '��Ϊ��ֵ
            glStatus = RegSetValueEx(llKeyID, psSubKey, 0&, REG_SZ, 0&, 0&)
        Else '��Ϊ����ֵ
            glStatus = RegSetValueEx(llKeyID, psSubKey, 0&, REG_SZ, ByVal psKeyValue, Len(psKeyValue) + 1)
        End If
        glStatus = RegCloseKey(llKeyID)
    End If

End Sub
Public Sub gpvSetKeyDwordValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal plKeyValue As Long)
    'ok
    '�� �ܣ�����ע����е�˫��ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' plKeyValue Long Ҫ����˫��ֵ
    ' ���: ��
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Sub
    End If
    
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then '�ɹ�������ֵ
        glStatus = RegSetValueEx(llKeyID, psSubKey, 0&, REG_DWORD, plKeyValue, Len(plKeyValue))
        glStatus = RegCloseKey(llKeyID)
    End If

End Sub
Public Sub gpvSetKeyBinaryValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal plKeyValue As Long)
    'ok
    '�� �ܣ�����ע����еĶ�����ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' plKeyValue Long Ҫ���õĶ�����ֵ
    ' ���: ��
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '�򿪼���ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ���ã��Ӽ�δ�������Ĭ��ֵ��
        glStatus = REGAGENT_NOKEY
        Exit Sub
    End If
    
    '���ȴ�����
    glStatus = RegOpenKey(plKey, psKey, llKeyID)
    
    If glStatus = ERROR_SUCCESS Then '�ɹ�������ֵ
        glStatus = RegSetValueEx(llKeyID, psSubKey, 0&, REG_BINARY, plKeyValue, Len(plKeyValue))
        glStatus = RegCloseKey(llKeyID)
    End If

End Sub
Public Function gflCreateKey(ByVal plKey As Long, ByVal psKey As String) As Long
    'ok
    '�� �ܣ�����ע����е�����
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' ���: gflCreateKey Long ����������ID
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '��ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ����
        glStatus = REGAGENT_NOKEY
        Exit Function
    End If
    
    '��������
    glStatus = RegCreateKey(plKey, psKey, llKeyID)
    If glStatus = ERROR_SUCCESS Then
        gflCreateKey = llKeyID
    End If

End Function
Public Sub gpvDeleteKey(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String)
    'ok
    '�� �ܣ�ɾ��ע����е�����
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' ���: ��
    ' Ӱ��: glStatus Long ״ֵ̬
    Dim llKeyID As Long '��ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ����
        glStatus = REGAGENT_NOKEY
        Exit Sub
    End If
    
    '���ô��������ж��Ƿ��������
    glStatus = RegCreateKey(plKey, psKey, llKeyID)
    glStatus = RegDeleteKey(llKeyID, ByVal psSubKey)

End Sub
Public Sub gpvDeleteKeyValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String)
    'ok
    '�� �ܣ�ɾ��ע����еļ�ֵ
    '�� ����
    ' ����: plKey Long ������
    ' psKey String ������
    ' psSubKey String �Ӽ���
    ' ���: ��
    ' Ӱ��: glStatus Long ״ֵ̬
    
    Dim llKeyID As Long '��ID
    
    glStatus = ERROR_SUCCESS '����ɹ�
    
    'ȷ��������Ч
    If Len(psKey) = 0 Then '����δ����
        glStatus = REGAGENT_NOKEY
        Exit Sub
    End If
    
    '���ô��������ж��Ƿ��������
    glStatus = RegCreateKey(plKey, psKey, llKeyID)
    glStatus = RegDeleteValue(llKeyID, ByVal psSubKey)

End Sub

Public Function GetSubKeys(ByVal psFullKey As String, ByRef psKeys() As String) As Boolean
    Dim plKey As Long, psKey As String
    If Not RegFullKeySplit(psFullKey, plKey, psKey, vbNull) Then GetSubKeys = False: Exit Function
    
    Dim trd As New clsTmsRegedit
    
    Dim dwIndex As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpcbClass As Long
    Dim lpClass As String
    Dim hKey As Long
    Dim lpftLastWriteTime As FILETIME
    Dim xx As Long
    xx = RegOpenKeyEx(plKey, psKey, 0, KEY_READ, hKey)
    If xx = 0 Then
        dwIndex = 0
        Do
            lpcbName = 255
            lpName = Space(lpcbName)
            lpcbClass = 255
            lpClass = Space(lpcbClass)
            
            xx = RegEnumKeyEx(hKey, ByVal dwIndex, ByVal lpName, lpcbName, 0, ByVal vbNullString, lpcbClass, lpftLastWriteTime)
            If xx = ERROR_NO_MORE_ITEMS Then Exit Do
            
            ReDim Preserve psKeys(0 To dwIndex)
            
            'On Error Resume Next
            psKeys(dwIndex) = Trim(lpName)
            
            dwIndex = dwIndex + 1
        Loop While xx <> trd.CONST_ERROR_NO_MORE_ITEMS
    Else
        GetSubKeys = False
        'IfStop = True
    End If
    RegCloseKey hKey
End Function
Public Sub GetSubValues(ByVal psFullKey As String, ByRef psKeys() As String)

End Sub
Public Function RegFullKeySplit(ByVal psFullKey As String, ByRef plKey As Long, ByRef psKey As String, ByRef psSubKey As Variant) As Boolean
    If 1 = 0 Then
err1:   RegFullKeySplit = False
        Exit Function
    End If
    RegFullKeySplit = True
    Select Case Left(psFullKey, InStr(psFullKey, "\") - 1)
    Case "HKEY_CLASSES_ROOT"
        plKey = HKEY_CLASSES_ROOT
    Case "HKCR"
        plKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_CONFIG"
        plKey = HKEY_CURRENT_CONFIG
    Case "HKCC"
        plKey = HKEY_CURRENT_CONFIG
    Case "HKEY_CURRENT_USER"
        plKey = HKEY_CURRENT_USER
    Case "HKCU"
        plKey = HKEY_CURRENT_USER
    Case "HKEY_DYN_DATA"
        plKey = HKEY_DYN_DATA
    Case "HKDD"
        plKey = HKEY_DYN_DATA
    Case "HKEY_LOCAL_MACHINE"
        plKey = HKEY_LOCAL_MACHINE
    Case "HKLM"
        plKey = HKEY_LOCAL_MACHINE
    Case "HKEY_PERFORMANCE_DATA"
        plKey = HKEY_PERFORMANCE_DATA
    Case "HKPD"
        plKey = HKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        plKey = HKEY_USERS
    Case "HKU"
        plKey = HKEY_USERS
    Case Else
        RegFullKeySplit = False
        Exit Function
    End Select
    
    If psSubKey = vbNull Then
        psKey = Right(psFullKey, Len(psFullKey) - InStr(psFullKey, "\"))
    Else
        psKey = Mid(psFullKey, InStr(psFullKey, "\") + 1, InStrRev(psFullKey, "\") - InStr(psFullKey, "\"))
        psSubKey = Right(psFullKey, Len(psFullKey) - InStrRev(psFullKey, "\"))
    End If
End Function

'�� �ܣ���ע�����ȡ�ô�ֵ
Function GetKeyStringValue(ByVal psFullKey As String) As String
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then GetKeyStringValue = gfsGetKeyStringValue(plKey, psKey, psSubKey)
End Function

'�� �ܣ���ע�����ȡ�ö�����ֵ
Function GetKeyBinaryValue(ByVal psFullKey As String) As Long
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then GetKeyBinaryValue = gflGetKeyBinaryValue(plKey, psKey, psSubKey)
End Function

'�� �ܣ���ע�����ȡ��˫��ֵ
Function GetKeyDwordValue(ByVal psFullKey As String) As Long
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then GetKeyDwordValue = gflGetKeyDwordValue(plKey, psKey, psSubKey)
End Function

'�� �ܣ�����ע����еĴ�ֵ
Sub SetKeyStringValue(ByVal psFullKey As String, ByVal psKeyValue As String)
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then gpvSetKeyStringValue plKey, psKey, psSubKey, psKeyValue
End Sub

'�� �ܣ�����ע����е�˫��ֵ
Sub SetKeyDwordValue(ByVal psFullKey As String, ByVal plKeyValue As Long)
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then gpvSetKeyDwordValue plKey, psKey, psSubKey, psKeyValue
End Sub

'�� �ܣ�����ע����еĶ�����ֵ
Sub SetKeyBinaryValue(ByVal psFullKey As String, ByVal plKeyValue As Long)
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then gpvSetKeyBinaryValue plKey, psKey, psSubKey, psKeyValue
End Sub

'�� �ܣ�����ע����е�����
Function CreateKey(ByVal psFullKey As String) As Long
    Dim plKey As Long, psKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, vbNull) Then CreateKey = gflCreateKey(plKey, psKey)
End Function

'�� �ܣ�ɾ��ע����е�����
Sub DeleteKey(ByVal psFullKey As String, ByVal psSubKey As String)
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then DeleteKey = gpvDeleteKey(plKey, psKey, psSubKey)
End Sub

'�� �ܣ�ɾ��ע����еļ�ֵ
Sub DeleteKeyValue(ByVal psFullKey As String, ByVal psSubKey As String)
    Dim plKey As Long, psKey As String, psSubKey As String
    If RegFullKeySplit(psFullKey, plKey, psKey, psSubKey) Then DeleteKeyValue = gpvDeleteKeyValue(plKey, psKey, psSubKey)
End Sub

'Function gfsGetKeyStringValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As String
'�� �ܣ���ע�����ȡ�ô�ֵ

'Function gflGetKeyBinaryValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As Long
'�� �ܣ���ע�����ȡ�ö�����ֵ

'Function gflGetKeyDwordValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String) As Long
'�� �ܣ���ע�����ȡ��˫��ֵ

'Sub gpvSetKeyStringValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal psKeyValue As String)
'�� �ܣ�����ע����еĴ�ֵ

'Sub gpvSetKeyDwordValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal plKeyValue As Long)
'�� �ܣ�����ע����е�˫��ֵ

'Sub gpvSetKeyBinaryValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String, ByVal plKeyValue As Long)
'�� �ܣ�����ע����еĶ�����ֵ

'Function gflCreateKey(ByVal plKey As Long, ByVal psKey As String) As Long
'�� �ܣ�����ע����е�����

'Sub gpvDeleteKey(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String)
'�� �ܣ�ɾ��ע����е�����

'Sub gpvDeleteKeyValue(ByVal plKey As Long, ByVal psKey As String, ByVal psSubKey As String)
'�� �ܣ�ɾ��ע����еļ�ֵ
