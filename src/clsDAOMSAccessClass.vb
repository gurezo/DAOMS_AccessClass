' @(h) clsDAOMSAccess.vb              ver 01.00.00
'
' @(s)
' このﾓｼﾞｭｰﾙはMicrosoft Access用のﾓｼﾞｭｰﾙです
' このﾓｼﾞｭｰﾙを使用する場合は、参照設定で、MicroSoft DAO3.6
' いずれかを選択していることを確認してください
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
    ' 機能　　 :MDBﾌｧｲﾙの切断処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - ｴﾗｰ番号
    '
    ' 引き数　 :ARG1 - ﾃﾞｰﾀﾍﾞｰｽ変数
    '
    ' 機能説明 :
    '
    ' 備考　　 :
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
    ' 機能　　 :MDBﾌｧｲﾙの接続処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - ｴﾗｰ番号
    '
    ' 引き数　 :ARG1 - ﾜｰｸｽﾍﾟｰｽ変数
    ' 　　　　  ARG2 - ﾃﾞｰﾀﾍﾞｰｽ変数
    ' 　　　　  ARG3 - ﾌﾙﾊﾟｽ付ﾃﾞｰﾀﾍﾞｰｽ名
    '
    ' 機能説明 :
    '
    ' 備考　　 :
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
    ' 機能　　 :DB最適化処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - ｴﾗｰ番号
    '
    ' 引き数　 :ARG1 - ﾒｯｾｰｼﾞ表示ﾌﾗｸﾞ
    ' 　　　　  ARG2 - ﾌﾙﾊﾟｽ付ﾃﾞｰﾀﾍﾞｰｽ名
    '
    ' 機能説明 :ｱｸｾｽMDBﾌｧｲﾙｻｲｽﾞ最適化処理
    '
    ' 備考　　 :
    '
    Function intMdbConpact(ByRef MsgFlg As Short, _
                           ByRef FPDbName As String) As Integer

        intMdbConpact = 0

        ''ﾒｯｾｰｼﾞﾌﾗｸﾞが0の時、ﾒｯｾｰｼﾞのYes/Noを表示
        If MsgFlg = 0 Then
            If MsgBox("データベースを最適化します。よろしいですか？", 292, "確認") = MsgBoxResult.No Then
                Exit Function
            End If
        End If

        ''最適化
        Try
            DBEng.CompactDatabase(FPDbName, FPDbName & "2")
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("最適化に失敗しました。", 48, "確認")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        ''最適化前ﾌｧｲﾙを削除
        Try
            Kill(FPDbName)
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("最適化元ファイル削除に失敗しました。", 48, "確認")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        ''名前を戻す（最適化を行う時は、最適化後のﾌｧｲﾙ名は最適化前と異なる）
        Try
            Rename(FPDbName & "2", FPDbName)
        Catch ex As Exception
            If MsgFlg = 0 Then
                Call MsgBox("最適化ファイル名変更に失敗しました。", 48, "確認")
            End If
            intMdbConpact = Err.Number
            Exit Function
        End Try

        If MsgFlg = 0 Then
            Call MsgBox("最適化を終了しました。", 48, "確認")
        End If

    End Function

    ' @(f)
    '
    ' 機能　　 :MdbSQL文実行処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - ｴﾗｰ番号
    '
    ' 引き数　 :ARG1 - ﾃﾞｰﾀﾍﾞｰｽ変数
    ' 　　　    ARG2 - SQL文
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Function intMdbExcute(ByRef UseDb As DAO.Database, _
                          ByRef SQLMoji As String) As Integer

        intMdbExcute = 0

        ''各種SQL文実行(Insert,UpDate)
        Try
            UseDb.Execute(SQLMoji)
        Catch ex As Exception
            intMdbExcute = Err.Number
        End Try

    End Function

    ' @(f)
    '
    ' 機能　　 :DAORecordset実行処理
    '
    ' 返り値　 :正常終了 - 0
    ' 　　　    ｴﾗｰ終了 - ｴﾗｰ番号
    '
    ' 引き数　 :ARG1 - ﾃﾞｰﾀﾍﾞｰｽ変数
    ' 　　　    ARG2 - ﾚｺｰﾄﾞｾｯﾄ変数
    ' 　　　    ARG3 - SQL文
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Function intMdbSelect(ByRef UseDb As DAO.Database, _
                          ByRef UseRec As DAO.Recordset, _
                          ByRef SQLMoji As String) As Integer

        intMdbSelect = 0

        ''ﾚｺｰﾄﾞ数初期化
        pMdbItem.MCount = 0

        ''ｾﾚｸﾄ実行
        Try
            UseRec = UseDb.OpenRecordset(SQLMoji, DAO.RecordsetTypeEnum.dbOpenSnapshot)
        Catch ex As Exception
            intMdbSelect = Err.Number
            Exit Function
        End Try

        ''ﾚｺｰﾄﾞ数格納
        pMdbItem.MCount = UseRec.RecordCount

    End Function

End Class
