VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'���C�u�������X�g�̐ݒ� (�ݒu�t�H���_�̓��[�N�u�b�N�Ɠ����f�B���N�g��)
Const FILENAME_LIBLIST As String = "libdef.txt" '���C�u�������X�g�̃t�@�C����
Const FILENAME_EXPORT As String = "ThisWorkbook-sjis.cls" '�G�N�X�|�[�g cls�t�@�C����
Const EXPORT_FOLDER As String = "src"

'�W�����W���[��
Const vbext_ct_StdModule As Integer = 1
'�N���X ���W���[��
Const vbext_ct_ClassModule As Integer = 2
'Microsoft Forms
Const vbext_ct_MSForm As Integer = 3
'ActiveX �f�U�C�i�[
Const vbext_ct_ActiveXDesigner As Integer = 11
'�h�L�������g ���W���[��
Const vbext_ct_Document As Integer = 100

'-------------------------------------------------------------------------------
' Name: reloadModule
' Func: �w�肵���t�@�C����reload FILENAME_LIBLIST �Ŏw��
' Reference from
'-------------------------------------------------------------------------------
' Text Scripting on VBA v1.0.0
' last update: 2013-01-03
' HATANO Hirokazu
'
' Detail: http://rsh.csh.sh/text-scripting-vba/
'  See Also: http://d.hatena.ne.jp/language_and_engineering/20090731/p1
'-------------------------------------------------------------------------------
Public Sub reloadModule()
Attribute reloadModule.VB_ProcData.VB_Invoke_Func = " \n14"
  '�蓮�����[�h�p Public�֐�
  
  Dim msgError As String
  msgError = loadModule("." & Application.PathSeparator & FILENAME_LIBLIST)
  
  If Len(msgError) > 0 Then
    MsgBox msgError
  End If
End Sub

Public Sub exportThisWorkbook()
  'ThisWorkbook �蓮export�p Public�֐�
  Call exportModule("ThisWorkbook", FILENAME_EXPORT)
End Sub

'----------------------------- main Subs/Functions ---------------

Private Function loadModule(ByVal pathConf As String) As String
  'Main: ���W���[�����X�g�t�@�C���ɏ����Ă���O�����C�u������ǂݍ��ށB

  '1. �S���W���[�����폜
  Dim isClear As Boolean
  isClear = clearModules
  
  If isClear = False Then
    loadModule = "Error: �W�����W���[���̑S�폜�Ɏ��s���܂����B"
    Exit Function
  End If
  
  
  '2. ���W���[�����X�g�t�@�C���̑��݊m�F
  ' 2.1. ���W���[�����X�g�t�@�C���̐�΃p�X���擾
  pathConf = absPath(pathConf)
  
  ' 2.2. ���݃`�F�b�N
  Dim isExistList As Boolean
  isExistList = checkExistFile(pathConf)
  
  If isExistList = False Then
    loadModule = "Error: ���C�u�������X�g" & pathConf & "�����݂��܂���B"
    Exit Function
  End If


  '3. ���W���[�����X�g�t�@�C���̓ǂݍ���&�z��
  Dim arrayModules As Variant
  arrayModules = list2array(pathConf)
  
  If UBound(arrayModules) = 0 Then
    loadModule = "Error: ���C�u�������X�g�ɗL���ȃ��W���[���̋L�q�����݂��܂���B"
    Exit Function
  End If

  
  '4. �e���W���[���t�@�C���ǂݍ���
  Dim i As Integer
  Dim msgError As String
  msgError = ""
  
  ' �z���0�n�܂�B(�ő�l: �z���-1)
  For i = 0 To UBound(arrayModules) - 1
    Dim pathModule As String
    pathModule = arrayModules(i)
    
    '4.1. ���W���[�����X�g�t�@�C���̑��݊m�F
    ' 4.1.1. ���W���[�����X�g�t�@�C���̐�΃p�X���擾
    pathModule = absPath(pathModule)
  
    ' 4.1.2. ���݃`�F�b�N
    Dim isExistModule As Boolean
    isExistModule = checkExistFile(pathModule)
  
    '4.2. ���W���[���ǂݍ���
    If isExistModule = True Then
      ThisWorkbook.VBProject.VBComponents.Import pathModule
    Else
      msgError = msgError & pathModule & " �͑��݂��܂���B" & vbCrLf
    End If
  Next i
  loadModule = msgError

End Function



'----------------------------- Functions / Subs ---------------

Private Sub exportModule(ByVal nameModule As String, ByVal nameFile As String)

  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
    
    If component.Name = nameModule Then
      component.Export ThisWorkbook.Path & Application.PathSeparator & nameFile
      MsgBox nameModule & " �� " & ThisWorkbook.Path & _
                  Application.PathSeparator & nameFile & " �Ƃ��ĕۑ����܂����B"
    End If
    
  Next component

End Sub




'----------------------------- common Functions / Subs ---------------
Private Function clearModules() As Boolean
  '�W�����W���[��/�N���X���W���[��������(�S�폜)
  
  Dim component As Object
  For Each component In ThisWorkbook.VBProject.VBComponents
      
    '�W�����W���[��(Type=1) / �N���X���W���[��(Type=2)��S�č폜
    If component.Type = 1 Or component.Type = 2 Then
      ThisWorkbook.VBProject.VBComponents.Remove component
    End If
    
  Next component
  
  '�W�����W���[��/�N���X���W���[���̍��v����0�ł����OK
  Dim cntBAS As Long
  cntBAS = countBAS()
  
  Dim cntClass As Long
  cntClass = countClasses()
        
  If cntBAS = 0 And cntClass = 0 Then
    clearModules = True
  Else
    clearModules = False
  End If

End Function



Private Function countBAS() As Long
  Dim count As Long
  count = countComponents(1) 'Type 1: bas
  countBAS = count
End Function



Private Function countClasses() As Long
  Dim count As Long
  count = countComponents(2) 'Type 2: class
  countClasses = count
End Function



Private Function countComponents(ByVal numType As Integer) As Long
  '���݂���W�����W���[��/�N���X���W���[���̐��𐔂���
  
  Dim i As Long
  Dim count As Long
  count = 0
  
  With ThisWorkbook.VBProject
    For i = 1 To .VBComponents.count
      If .VBComponents(i).Type = numType Then
        count = count + 1
      End If
    Next i
  End With

  countComponents = count
End Function



Private Function absPath(ByVal pathFile As String) As String
  ' �t�@�C���p�X���΃p�X�ɕϊ�
  
  Dim nameOS As String
  nameOS = Application.OperatingSystem
  
  'replace Win backslash(Chr(92))
  pathFile = Replace(pathFile, Chr(92), Application.PathSeparator)
  
  'replace Mac ":"Chr(58)
  pathFile = Replace(pathFile, ":", Application.PathSeparator)
  
  'replace Unix "/"Chr(47)
  pathFile = Replace(pathFile, "/", Application.PathSeparator)


  Select Case Left(pathFile, 1)
  
    'Case1. . �Ŏn�܂�ꍇ(���Ύw��)
    Case ".":
  
      Select Case Left(pathFile, 2)
        
        ' Case1-1. ���Ύw�� "../" �Ή�
        Case "..":
          'MsgBox "Case1-1: " & pathFile
          absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
          Exit Function
    
        ' Case1-2. ���Ύw�� "./" �Ή�
        Case Else:
          'MsgBox "Case1-2: " & pathFile
          absPath = ThisWorkbook.Path & Mid(pathFile, 2, Len(pathFile) - 1)
          Exit Function
    
      End Select
    
    'Case2. ��؂蕶���Ŏn�܂�ꍇ (��Ύw��)
    Case Application.PathSeparator:
    
      ' Case2-1. Windows Network Drive ( chr(92) & chr(92) & "hoge")
      'MsgBox "Case2-1: " & pathFile
      If Left(pathFile, 2) = Chr(92) & Chr(92) Then
        absPath = pathFile
        Exit Function
      
      Else
      ' Case2-2. Mac/UNIX Absolute path (/hoge)
        absPath = pathFile
        Exit Function
      
      End If
    
  End Select


  'Case3. [A-z][0-9]�Ŏn�܂�ꍇ
  ' (Mac��Office�Ő��K�\�����g����� select���ɓ����ׂ�...)

  ' Case3-1.�h���C�u���^�[�Ή�("c:" & chr(92) �� "c" &
  ' chr(92) & chr(92)�ɂȂ��Ă��܂��̂ŏ����߂�)
  If nameOS Like "Windows *" And Left(pathFile, 2) Like "[A-z]" & _
                                                    Application.PathSeparator Then
    'MsgBox "Case3-1" & pathFile
    absPath = Replace(pathFile, Application.PathSeparator, ":", 1, 1)
    Exit Function
  End If
 
  ' Case3-2. ���w�� "filename"�Ή�
  If Left(pathFile, 1) Like "[0-9]" Or Left(pathFile, 1) Like "[A-z]" Then
    absPath = ThisWorkbook.Path & Application.PathSeparator & pathFile
    Exit Function
  Else
    MsgBox "Error[AbsPath]: fail to get absolute path."
  
  End If

End Function




Private Function checkExistFile(ByVal pathFile As String) As Boolean

  On Error GoTo Err_dir
  If Dir(pathFile) = "" Then
    checkExistFile = False
  Else
    checkExistFile = True
  End If

  Exit Function

Err_dir:
  checkExistFile = False

End Function

'���X�g�t�@�C����z��ŕԂ�(�s����'(�R�����g)�̍s & ��s�͖�������)
Private Function list2array(ByVal pathFile As String) As Variant
    
  Dim nameOS As String
  nameOS = Application.OperatingSystem
        
  '1. ���X�g�t�@�C���̓ǂݎ��
  Dim fp As Integer
  fp = FreeFile
  Open pathFile For Input As #fp
  
  '2. ���X�g�̔z��
  Dim arrayOutput() As String
  Dim countLine As Integer
  countLine = 0
  ReDim Preserve arrayOutput(countLine) ' �z��0�ŕԂ��ꍇ�����邽��
  
  Do Until EOF(fp)
    '���C�u�������X�g��1�s������
    Dim strLine As String
    Line Input #fp, strLine

    Dim isLf As Long
    isLf = InStr(strLine, vbLf)
    
    If nameOS Like "Windows *" And Not isLf = 0 Then
      'OS��Windows ���� ���X�g�� LF���܂܂��ꍇ (�t�@�C����UNIX�`��)
      '�t�@�C���S�̂�1�s�Ɍ����Ă��܂��B
      
      Dim arrayLineLF As Variant
      arrayLineLF = Split(strLine, vbLf)
    
      Dim i As Integer
      For i = 0 To UBound(arrayLineLF) - 1
        '�s���� '(�R�����g) �ł͂Ȃ� & ��s�ł͂Ȃ��ꍇ
        If Not Left(arrayLineLF(i), 1) = "'" And Len(arrayLineLF(i)) > 0 Then
      
          '�z��ւ̒ǉ�
          countLine = countLine + 1
          ReDim Preserve arrayOutput(countLine)
          arrayOutput(countLine - 1) = arrayLineLF(i)
        End If
      Next i
              
    
    Else
      'OS��Windows and �t�@�C����Windows�`�� (�ϊ��s�v)
      'OS��MacOS X and �t�@�C����UNIX�`�� (�ϊ��s�v)
      
      'OS��MacOS X and �t�@�C����Windows�`��
      ' vbCr�����W���[���t�@�C�����𔭌��ł��Ȃ��Ȃ�B
      strLine = Replace(strLine, vbCr, "")
    
  
      '�s���� '(�R�����g) �ł͂Ȃ� & ��s�ł͂Ȃ��ꍇ
      If Not Left(strLine, 1) = "'" And Len(strLine) > 0 Then
      
        '�z��ւ̒ǉ�
        countLine = countLine + 1
        ReDim Preserve arrayOutput(countLine)
        arrayOutput(countLine - 1) = strLine
      End If
    
    End If
  Loop

  '3. ���X�g�t�@�C�������
  Close #fp
  
  '4. �߂�l��z��ŕԂ�
  list2array = arrayOutput
End Function

'-------------------------------------------------------------
' Name: ExportAll
' Func: ���ׂ�Export
' Reference from
' http://d.hatena.ne.jp/jamzz/20131002/1380696685
'-------------------------------------------------------------
Public Sub ExportAll()
    Dim export_path As String
    Dim full_path As String
    Dim vb_component As Object
    
    export_path = EXPORT_FOLDER
    
    For Each vb_component In ThisWorkbook.VBProject.VBComponents

        Dim extention As String
        Select Case vb_component.Type
            Case vbext_ct_StdModule
            '�W�����W���[��
            extention = ".bas"
            Case vbext_ct_ClassModule
            '�N���X ���W���[��
            extention = ".cls"
            Case vbext_ct_MSForm
            'Microsoft Forms
            extention = ".frm"
            Case vbext_ct_ActiveXDesigner
            'ActiveX �f�U�C�i�[
            extention = ".cls"
            Case vbext_ct_Document
            '�h�L�������g ���W���[��
             extention = ".cls"
        End Select
        
        ' �G�N�X�|�[�g
        full_path = ThisWorkbook.Path & "\\" & export_path & "\\" & vb_component.Name & extention
        Debug.Print "Export to " & full_path
        vb_component.Export full_path
    Next
End Sub

'-------------------------------------------------------------------------------
' Name: RunAllTests
' Func: �e�X�g���s
'-------------------------------------------------------------------------------
Public Sub RunAllTests()
  Dim Runner As New TestRunner
  Runner.Run
  Set Runner = Nothing
End Sub