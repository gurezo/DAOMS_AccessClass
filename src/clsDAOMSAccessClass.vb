' @(h) clsDAOMSAccess.vb              ver 01.00.00
'
' @(s)
' ����Ӽޭ�ق�Microsoft Access�p��Ӽޭ�قł�
' ����Ӽޭ�ق��g�p����ꍇ�́A�Q�Ɛݒ�ŁAMicroSoft DAO3.6
' �����ꂩ��I�����Ă��邱�Ƃ��m�F���Ă�������
'
Option Strict Off
Option Explicit On 

Public Class DAOAccess

    Private DBEng As New DAO.DBEngine

    Structure MdbItem
        Dim MWorkspace As DAO.Workspace
        Dim MDataBase As DAO.Database
        Dim MRecord As DAO.Recordset
        Dim MCount As Integer
    End Structure
    Public pMdbItem As MdbItem

    ' @(f)
    '
    ' �@�\�@�@ :MDB̧�ق̐ؒf����
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �װ�ԍ�
    '
    ' �������@ :ARG1 - �ް��ް��ϐ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    '
    Function intMdbClose(ByRef UseDb As DAO.Database) As Integer

        intMdbClose = 0

        Try
            UseDb.Close()
        Catch ex As Exception
            intMdbClose = Err.Number
        End Try

    End Function


    ' @(f)
    '
    ' �@�\�@�@ :MDB̧�ق̐ڑ�����
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �װ�ԍ�
    '
    ' �������@ :ARG1 - ܰ���߰��ϐ�
    ' �@�@�@�@  ARG2 - �ް��ް��ϐ�
    ' �@�@�@�@  ARG3 - ���߽�t�ް��ް���
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    '
    Function intMdbOpen(ByRef UseWks As DAO.Workspace, _
                        ByRef UseDb As DAO.Database, _
                        ByRef FPDbName As String) As Integer

        intMdbOpen = 0

        Try
            UseWks = DBEng.CreateWorkspace("", "admin", "", DAO.WorkspaceTypeEnum.dbUseJet)
        Catch ex As Exception
            intMdbOpen = Err.Number
        End Try

        Try
            UseDb = UseWks.OpenDatabase(FPDbName, False)
        Catch ex As Exception
            intMdbOpen = Err.Number
        End Try

    End Function



    ' @(f)
    '
    ' �@�\�@�@ :DB�œK������
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �װ�ԍ�
    '
    ' �������@ :ARG1 - ү���ޕ\���׸�
    ' �@�@�@�@  ARG2 - ���߽�t�ް��ް���
    '
    ' �@�\���� :����MDB̧�ٻ��ލœK������
    '
    ' ���l�@�@ :
    '
    Function intMdbConpact(ByRef MsgFlg As Short, _
                           ByRef FPDbName As String) As Integer

        intMdbConpact = 0

        ''ү�����׸ނ�0�̎��Aү���ނ�Yes/No��\��
        If MsgFlg = 0 Then
            If MsgBox("�f�[�^�x�[�X���œK�����܂��B��낵���ł����H", 292, "�m�F") = MsgBoxResult.No Then
                Exit Function
            End If
        End If

        ''�œK��
        Try
            DBEng.CompactDatabase(FPDbName, FPDbName & "2")
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("�œK���Ɏ��s���܂����B", 48, "�m�F")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        ''�œK���O̧�ق��폜
        Try
            Kill(FPDbName)
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("�œK�����t�@�C���폜�Ɏ��s���܂����B", 48, "�m�F")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        ''���O��߂��i�œK�����s�����́A�œK�����̧�ٖ��͍œK���O�ƈقȂ�j
        Try
            Rename(FPDbName & "2", FPDbName)
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("�œK���t�@�C�����ύX�Ɏ��s���܂����B", 48, "�m�F")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        If MsgFlg = 0 Then
            Call MsgBox("�œK�����I�����܂����B", 48, "�m�F")
        End If

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :MdbSQL�����s����
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �װ�ԍ�
    '
    ' �������@ :ARG1 - �ް��ް��ϐ�
    ' �@�@�@    ARG2 - SQL��
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Function intMdbExcute(ByRef UseDb As DAO.Database, _
                          ByRef SQLMoji As String) As Integer

        intMdbExcute = 0

        ''�e��SQL�����s(Insert,UpDate)
        Try
            UseDb.Execute(SQLMoji)
        Catch ex As Exception
            intMdbExcute = Err.Number
        End Try

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :DAORecordset���s����
    '
    ' �Ԃ�l�@ :����I�� - 0
    ' �@�@�@    �װ�I�� - �װ�ԍ�
    '
    ' �������@ :ARG1 - �ް��ް��ϐ�
    ' �@�@�@    ARG2 - ں��޾�ĕϐ�
    ' �@�@�@    ARG3 - SQL��
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Function intMdbSelect(ByRef UseDb As DAO.Database, _
                          ByRef UseRec As DAO.Recordset, _
                          ByRef SQLMoji As String) As Integer

        intMdbSelect = 0

        ''ں��ސ�������
        pMdbItem.MCount = 0

        ''�ڸĎ��s
        Try
            UseRec = UseDb.OpenRecordset(SQLMoji, DAO.RecordsetTypeEnum.dbOpenSnapshot)
        Catch ex As Exception
            intMdbSelect = Err.Number
            Exit Function
        End Try

        ''ں��ސ��i�[
        pMdbItem.MCount = UseRec.RecordCount

    End Function

End Class
