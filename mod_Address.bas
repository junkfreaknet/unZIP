Attribute VB_Name = "mod_Address"
Option Compare Database
Option Explicit

    Const CS_TABLE_NAME_新旧変換テーブル As String = "TO_新旧変換テーブル"
    
    Const CS_FIELD_NAME_旧ID As String = "旧ID"
    Const CS_FIELD_NAME_新ID As String = "新ID"
    
    Const CS_TABLE_NAME_TO_浜松市旧北区 As String = "TO_浜松市旧北区"
    
    Const CS_FIELD_NAME_町名 As String = "町名"
    Const CS_FIELD_NAME_新標準地域コード As String = "新標準地域コード"
    
Function getチェックデジット都道府県コード(in_Code As String) As String

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
    
    getチェックデジット都道府県コード = rv
    
End Function

Function get地域公共団体コード6桁(in_Code_5 As String) As String
    
    Dim rv As String
    'Debug.Print in_Code_5
    
    get地域公共団体コード6桁 = in_Code_5 & getチェックデジット都道府県コード(in_Code_5)
    
End Function

Function get_境界線コード_補正済み(in_Code As String, in_Place_Name As String) As String

    '*****浜松市の区の変換も行う。
    Dim rv
    
    Dim LengthOF_Code As Integer
    Dim LengthOf_Code_New As Integer
    
    Dim codeMae As String
    Dim codeAto As String
    Const CS_LenAreaCode As Integer = 5
    
    rv = CS_SPACE
    LengthOF_Code = Len(in_Code)
    codeMae = get_市旧新区変換_浜松市_20240101(Mid(in_Code, 1, CS_LenAreaCode), in_Place_Name) '*****浜松市20240101
    codeAto = Mid(in_Code, CS_LenAreaCode + 1, LengthOF_Code - CS_LenAreaCode)

    rv = get地域公共団体コード6桁(codeMae) + codeAto
    
    get_境界線コード_補正済み = rv
    
End Function

Function get_市旧新区変換_浜松市_20240101(in_Code As String, in_Place_Name As String) As String

    '*****  引数は5桁で
    
    Dim rv As String
    
    Dim dbs As DAO.Database
    Dim rst_In As DAO.Recordset
    Dim SQLString As String
    Dim rst_KitaKu As DAO.Recordset
    
    Const CS_CODE_旧区_北区 As String = "22135"
    
    rv = CS_SPACE
    
    Set dbs = CurrentDb()
    SQLString = "select * from " & CS_TABLE_NAME_新旧変換テーブル & _
                " where " & _
                    CS_FIELD_NAME_旧ID & "='" & in_Code & "'"
                    
    Set rst_In = dbs.OpenRecordset(SQLString)
    
    If rst_In.EOF Then
        rv = in_Code
    Else
        rst_In.MoveFirst
        rv = CStr(rst_In!新ID)
        '*****旧北区なら更にチェック
        If rv = CS_CODE_旧区_北区 Then
            SQLString = "select * from " & CS_TABLE_NAME_TO_浜松市旧北区 & " where " & CS_FIELD_NAME_町名 & "='" & in_Place_Name & "'"
            Set rst_KitaKu = dbs.OpenRecordset(SQLString)
            If rst_KitaKu.EOF Then
            Else
                rst_KitaKu.MoveFirst
                rv = CStr(rst_KitaKu!新標準地域コード)
            End If
            rst_KitaKu.Close
        End If
    End If
    
    rst_In.Close
    Set dbs = Nothing
    
    get_市旧新区変換_浜松市_20240101 = rv
    
End Function
