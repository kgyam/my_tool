Attribute VB_Name = "模块2"
Sub MYSQL建表语句()
    Dim sqlStr, sqlTemp, pkStr, columnName, columnType, columnLength, notExistTable, isNull, comment As String
    Dim tableListIndex, columnListIndex As Integer
    
    Dim avFilePath As String '文件路径
    Dim lvIntFileNum As Integer '空文件号
    Dim lvContents As String
    avFilePath = "D:\SQLServer.txt" '文件路径
    
    lvIntFileNum = FreeFile() '获取一个空文件
    
        '生成建表ddl
        Dim sql As String
        Dim tableName As String
        Dim tableCommStr As String
        Dim prikeyStr As String
        
        For si = 2 To Workbooks(1).Sheets.Count '从第三张表开始,遍历每一张表
            
            Dim isPk As String
            prikeyStr = ""
            tableCommStr = ""
            Set mysheet = Workbooks(1).Sheets(si) '表
            tableName = mysheet.Range("B1").Value '英文表名
            tableCommStr = mysheet.Range("B3").Value '数据表注释
            '如果数据库中已存在表，则删除表
            'if exists (select * from sysobjects where name='tableName')
            'sql = sql & "if exists (select * from sysobjects where name='" & tableName & "') " & vbCrLf
            
            'drop table tableName
            'sql = sql & "   drop table [" & tableName & "] " & vbCrLf
            'sql = sql & "go " & vbCrLf
              
            '开始创建表
            'create table tableName (
            sql = sql & "create table " & Trim(tableName) & "  ( " & vbCrLf
            
              
            For i = 5 To mysheet.UsedRange.Rows.Count '从第五行开始遍历所有的列
                Dim nameStr As String
                Dim typeStr As String
                Dim commStr As String
               
                
                
                 nameStr = mysheet.Range("B" & i).Value '字段名
                 typeStr = mysheet.Range("C" & i).Value '数据类型
                 commStr = mysheet.Range("F" & i).Value '中文注释
                
                If typeStr = "int" Or typeStr = "bigint" Or typeStr = "datetime" Or typeStr = "text" Or typeStr = "image" Or typeStr = "tinyint" Then
                
                   typeStr = Trim(mysheet.Range("C" & i).Value) '数据类型
                Else
                   typeStr = Trim(mysheet.Range("C" & i).Value) + "(" + Trim(Str(mysheet.Range("D" & i).Value)) + ")" '数据类型
                End If
                
               
                
                
                isNull = Trim(mysheet.Range("E" & i).Value) '是否为空
                isPk = Trim(mysheet.Range("G" & i).Value) '是否主键
                
                sql = sql & "   " & nameStr & " " & typeStr
                
                If isNull = "否" Then
                    sql = sql + "  NOT NULL"
                End If
                
                If Len(commStr) > 0 Then
                    sql = sql + " COMMENT '" + commStr + "'"
                End If
                
                If isPk = "是" Then
                     'sql = sql + " PRIMARY KEY"
                    prikeyStr = mysheet.Range("B" & i).Value
                End If
                
                  
                
                
               ' If Len(commStr) > 0 Then
               ' comm = comm & vbCrLf & "EXEC sp_addextendedproperty 'MS_Description', N'" & commStr & "','SCHEMA', N'dbo','TABLE', N'" & tableName & "','COLUMN', N'" & nameStr & "';"
               ' End If
            
            If i < mysheet.UsedRange.Rows.Count Then
                sql = sql & "," & vbCrLf '添加‘,’号
            End If
           
        Next i
        
        If Len(prikeyStr) > 0 Then
        sql = sql & "," & vbCrLf & "PRIMARY KEY(" & prikeyStr & ")" & vbCrLf
        Else
        sql = sql & vbCrLf
        End If
        
        
        sql = sql & ")"
        If Len(tableCommStr) > 0 Then
        sql = sql & " COMMENT='" & tableCommStr & "'" & vbCrLf
        End If
        
        
        sql = sql & "ENGINE=InnoDB DEFAULT CHARSET=utf8;" & vbCrLf & vbCrLf & vbCrLf
        
        ' sql = sql & "go" & vbCrLf
        ' sql = sql & comm & vbCrLf & "-----Create table " & tableName & " end." & vbCrLf & vbCrLf
        
        
     Next si
     ' 单张建表语句完成
     lvContents = lvContents + sql
    Open avFilePath For Output As #lvIntFileNum '打开处理文件
     Print #lvIntFileNum, (lvContents) '写文件
     Close #lvIntFileNum '关闭文件
     
     If notExistTable <> "" Then
        MsgBox "表" + notExistTable + "不存在！！"
     End If
     
     MsgBox "文件生成成功！文件路径" + avFilePath
    
End Sub






