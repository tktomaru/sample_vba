Attribute VB_Name = "ModuleFileBinary"
' ����m�F�p�̊֐�
Public Sub PT_WriteBinaryFile()
   Dim outputFileName As String
   Dim result As String
   ' �G�N�Z���̂���t�H���_��output.dat��ǂݍ���
   outputFileName = ThisWorkbook.Path & "\" & "output.dat"
   result = WriteBinaryFile(outputFileName, "010203040506070809101112AABBCCDDEEFF")
   ' �������݌��ʂ̓t�@�C�����Q�Ƃ��邱��
End Sub

' ����m�F�p�̊֐�
Public Sub PT_ReadBinaryFile()
   Dim outputFileName As String
   Dim outputHexText As String
   ' �G�N�Z���̂���t�H���_��output.dat��ǂݍ���
   outputFileName = ThisWorkbook.Path & "\" & "output.dat"
   outputHexText = ReadBinaryFile(outputFileName)
   ' ���b�Z�[�W�{�b�N�X�ɓǂݍ��݌��ʂ�\��
   MsgBox outputHexText
End Sub


' ���L���Q�l�ɂ��Ď�������
' �yVBA�zOpen�X�e�[�g�����g�Ńo�C�i���t�@�C����ǂݏ������� | �₳�����v���O���~���O���Y�^
' http://pg-sample.sagami-ss.net/?eid=9
'********************************************
' �o�C�i���f�[�^���e�X�g�t�@�C���ɏo��
' param strfil ���̓t�@�C����
' param strHexText 16�i���̃e�L�X�g�i��F"010203FF"�j
'
' return ����="1" ���s="0"
'********************************************
Function WriteBinaryFile(ByVal strfil As String, strHexText As String) As String
    '//�o�C�i���t�@�C����1�o�C�g���̓��o�͂ɂ�Byte�^��p����
    Dim buff() As Byte
    Dim i As Integer
    Dim fp As Long
    Dim outputLen As Integer
    
    ' �o�̓T�C�Y
    outputLen = (Len(strHexText) / 2) - 1
    ' �o�͗̈�
    ReDim buff(0 To outputLen) As Byte
    '�������݃f�[�^���Z�b�g
    For i = 0 To outputLen
        buff(i) = HexTextToDec(Mid(strHexText, i * 2 + 1, 2))
    Next

    ' FreeFile�֐��Ŏg�p�\�ȃt�@�C���ԍ������蓖��
    fp = FreeFile

    ' �t�@�C�������݂���ꍇ�͎w��A�h���X���㏑������邾���̂���
    ' �������ݑO�Ƀt�@�C�����폜���邩���g����U�N���A����
    Open strfil For Output As #fp
    Close (fp)

    ' �t�@�C���I�[�v��(�o�C�i���������݂ŃI�[�v���A�t�@�C�������݂��Ȃ��ꍇ�͐V�K�쐬)
    ' Open�X�e�[�g�����g��p���ăt�@�C���̓��o�͂��s���܂�
    ' ���[�h�ɉ��L�̂����ꂩ���w�肳��Ă���΃t�@�C�������݂��Ȃ��ꍇ�A�V�K�쐬����܂�
    ' �ǉ����[�h(Append)�A�o�C�i�����[�h(Binary)�A�o�̓��[�h(Output)�A�����_���A�N�Z�X���[�h(Random)
    ' ��https://msdn.microsoft.com/ja-jp/library/office/gg264163.aspx
    Open strfil For Binary Access Write As #fp

       '//�t�@�C���ɏ�������(�t�@�C���擪����̏������݂𖾎�)
       Put #fp, 1, buff

    '//�t�@�C�������
    Close (fp)
    ' ����
    WriteBinaryFile = "1"
End Function

 

'********************************************
'�e�X�g�t�@�C������o�C�i���f�[�^��ǂݍ���
' param strfil ���̓t�@�C����
'
' return ReadBinaryFile 16�i���̃e�L�X�g�i��F"010203FF"�j
'********************************************
Function ReadBinaryFile(ByVal strfil As String) As String
    Dim buff() As Byte
    Dim fp As Long
    Dim filesize As Long, NowLoc As Long
    Dim idx As Long, gyo As Long
    Dim strBinary As String
    Dim strOutput As String

    '//FreeFile�֐��Ŏg�p�\�ȃt�@�C���ԍ������蓖��
    fp = FreeFile

    '//�t�@�C�����J��
    Open strfil For Binary As #fp

    '//�t�@�C���T�C�Y���̓ǂݍ��ݗ̈���m�ۂ��ēǂݍ��ޏꍇ�̎�����
    'ReDim buff(FileLen(strfil))
    'Get #fp, 1, buff
    '//�����Ⴑ���܂�

    ' �o�͏�����
    strOutput = ""
    '//�t�@�C���̏I�[�܂Ŏw��T�C�Y(�ő�16�o�C�g)�J��Ԃ��ǂݍ���
    Do While NowLoc < LOF(fp)

        '//�ő�16�o�C�g���̗̈���m�ۂ�������
        If (LOF(fp) - NowLoc) >= 16 Then
            '//�c��̃t�@�C���T�C�Y��16�o�C�g�ȏ�̂Ƃ�
            ReDim buff(15)
        Else
            '//�ŏI�ǂݍ��ݎ�(497�o�C�g�`500�o�C�g��)�͎c��̃t�@�C���T�C�Y��16����
            ReDim buff(LOF(fp) - NowLoc - 1)
        End If

        '//�f�[�^��ǂݍ���
        Get #fp, , buff

        '//���݈ʒu����ێ�����(���[�vBreak����p)
        NowLoc = Loc(fp)

        '//�o�͕�����𐶐�
        For idx = 0 To UBound(buff)
            strBinary = strBinary + Right("00" & Hex(buff(idx)), 2)
        Next

        '//�V�[�g��1��ڂɌ��ʂ�\��
        strOutput = strOutput + strBinary
        gyo = gyo + 1
    Loop

    '//�t�@�C�������
    Close (fp)
    ReadBinaryFile = strOutput
End Function
