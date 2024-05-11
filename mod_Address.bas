Attribute VB_Name = "mod_Address"
Option Compare Database
Option Explicit

    Const CS_TABLE_NAME_�V���ϊ��e�[�u�� As String = "TO_�V���ϊ��e�[�u��"
    
    Const CS_FIELD_NAME_��ID As String = "��ID"
    Const CS_FIELD_NAME_�VID As String = "�VID"
    
    Const CS_TABLE_NAME_TO_�l���s���k�� As String = "TO_�l���s���k��"
    
    Const CS_FIELD_NAME_���� As String = "����"
    Const CS_FIELD_NAME_�V�W���n��R�[�h As String = "�V�W���n��R�[�h"
    
Function get�`�F�b�N�f�W�b�g�s���{���R�[�h(in_Code As String) As String

    Dim rv As String
    Dim rv_Int As Integer
    
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim e As Integer
    
    
    rv = ""
    a = CInt(Mid(in_Code, 1, 1))
    b = CInt(Mid(in_Code, 2, 1))
    c = CInt(Mid(in_Code, 3, 1))
    d = CInt(Mid(in_Code, 4, 1))
    e = CInt(Mid(in_Code, 5, 1))
    
    rv = CStr(11 - ((6 * a + 5 * b + 4 * c + 3 * d + 2 * e) Mod 11)) Mod 10
    
    get�`�F�b�N�f�W�b�g�s���{���R�[�h = rv
    
End Function

Function get�n������c�̃R�[�h6��(in_Code_5 As String) As String
    
    Dim rv As String
    'Debug.Print in_Code_5
    
    get�n������c�̃R�[�h6�� = in_Code_5 & get�`�F�b�N�f�W�b�g�s���{���R�[�h(in_Code_5)
    
End Function

Function get_���E���R�[�h_�␳�ς�(in_Code As String, in_Place_Name As String) As String

    '*****�l���s�̋�̕ϊ����s���B
    Dim rv
    
    Dim LengthOF_Code As Integer
    Dim LengthOf_Code_New As Integer
    
    Dim codeMae As String
    Dim codeAto As String
    Const CS_LenAreaCode As Integer = 5
    
    rv = CS_SPACE
    LengthOF_Code = Len(in_Code)
    codeMae = get_�s���V��ϊ�_�l���s_20240101(Mid(in_Code, 1, CS_LenAreaCode), in_Place_Name) '*****�l���s20240101
    codeAto = Mid(in_Code, CS_LenAreaCode + 1, LengthOF_Code - CS_LenAreaCode)

    rv = get�n������c�̃R�[�h6��(codeMae) + codeAto
    
    get_���E���R�[�h_�␳�ς� = rv
    
End Function

Function get_�s���V��ϊ�_�l���s_20240101(in_Code As String, in_Place_Name As String) As String

    '*****  ������5����
    
    Dim rv As String
    
    Dim dbs As DAO.Database
    Dim rst_In As DAO.Recordset
    Dim SQLString As String
    Dim rst_KitaKu As DAO.Recordset
    
    Const CS_CODE_����_�k�� As String = "22135"
    
    rv = CS_SPACE
    
    Set dbs = CurrentDb()
    SQLString = "select * from " & CS_TABLE_NAME_�V���ϊ��e�[�u�� & _
                " where " & _
                    CS_FIELD_NAME_��ID & "='" & in_Code & "'"
                    
    Set rst_In = dbs.OpenRecordset(SQLString)
    
    If rst_In.EOF Then
        rv = in_Code
    Else
        rst_In.MoveFirst
        rv = CStr(rst_In!�VID)
        '*****���k��Ȃ�X�Ƀ`�F�b�N
        If rv = CS_CODE_����_�k�� Then
            SQLString = "select * from " & CS_TABLE_NAME_TO_�l���s���k�� & " where " & CS_FIELD_NAME_���� & "='" & in_Place_Name & "'"
            Set rst_KitaKu = dbs.OpenRecordset(SQLString)
            If rst_KitaKu.EOF Then
            Else
                rst_KitaKu.MoveFirst
                rv = CStr(rst_KitaKu!�V�W���n��R�[�h)
            End If
            rst_KitaKu.Close
        End If
    End If
    
    rst_In.Close
    Set dbs = Nothing
    
    get_�s���V��ϊ�_�l���s_20240101 = rv
    
End Function
