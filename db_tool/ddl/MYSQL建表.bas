Attribute VB_Name = "ģ��2"
Sub MYSQL�������()
    Dim sqlStr, sqlTemp, pkStr, columnName, columnType, columnLength, notExistTable, isNull, comment As String
    Dim tableListIndex, columnListIndex As Integer
    
    Dim avFilePath As String '�ļ�·��
    Dim lvIntFileNum As Integer '���ļ���
    Dim lvContents As String
    avFilePath = "D:\SQLServer.txt" '�ļ�·��
    
    lvIntFileNum = FreeFile() '��ȡһ�����ļ�
    
        '���ɽ���ddl
        Dim sql As String
        Dim tableName As String
        Dim tableCommStr As String
        Dim prikeyStr As String
        
        For si = 2 To Workbooks(1).Sheets.Count '�ӵ����ű�ʼ,����ÿһ�ű�
            
            Dim isPk As String
            prikeyStr = ""
            tableCommStr = ""
            Set mysheet = Workbooks(1).Sheets(si) '��
            tableName = mysheet.Range("B1").Value 'Ӣ�ı���
            tableCommStr = mysheet.Range("B3").Value '���ݱ�ע��
            '������ݿ����Ѵ��ڱ���ɾ����
            'if exists (select * from sysobjects where name='tableName')
            'sql = sql & "if exists (select * from sysobjects where name='" & tableName & "') " & vbCrLf
            
            'drop table tableName
            'sql = sql & "   drop table [" & tableName & "] " & vbCrLf
            'sql = sql & "go " & vbCrLf
              
            '��ʼ������
            'create table tableName (
            sql = sql & "create table " & Trim(tableName) & "  ( " & vbCrLf
            
              
            For i = 5 To mysheet.UsedRange.Rows.Count '�ӵ����п�ʼ�������е���
                Dim nameStr As String
                Dim typeStr As String
                Dim commStr As String
               
                
                
                 nameStr = mysheet.Range("B" & i).Value '�ֶ���
                 typeStr = mysheet.Range("C" & i).Value '��������
                 commStr = mysheet.Range("F" & i).Value '����ע��
                
                If typeStr = "int" Or typeStr = "bigint" Or typeStr = "datetime" Or typeStr = "text" Or typeStr = "image" Or typeStr = "tinyint" Then
                
                   typeStr = Trim(mysheet.Range("C" & i).Value) '��������
                Else
                   typeStr = Trim(mysheet.Range("C" & i).Value) + "(" + Trim(Str(mysheet.Range("D" & i).Value)) + ")" '��������
                End If
                
               
                
                
                isNull = Trim(mysheet.Range("E" & i).Value) '�Ƿ�Ϊ��
                isPk = Trim(mysheet.Range("G" & i).Value) '�Ƿ�����
                
                sql = sql & "   " & nameStr & " " & typeStr
                
                If isNull = "��" Then
                    sql = sql + "  NOT NULL"
                End If
                
                If Len(commStr) > 0 Then
                    sql = sql + " COMMENT '" + commStr + "'"
                End If
                
                If isPk = "��" Then
                     'sql = sql + " PRIMARY KEY"
                    prikeyStr = mysheet.Range("B" & i).Value
                End If
                
                  
                
                
               ' If Len(commStr) > 0 Then
               ' comm = comm & vbCrLf & "EXEC sp_addextendedproperty 'MS_Description', N'" & commStr & "','SCHEMA', N'dbo','TABLE', N'" & tableName & "','COLUMN', N'" & nameStr & "';"
               ' End If
            
            If i < mysheet.UsedRange.Rows.Count Then
                sql = sql & "," & vbCrLf '��ӡ�,����
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
     ' ���Ž���������
     lvContents = lvContents + sql
    Open avFilePath For Output As #lvIntFileNum '�򿪴����ļ�
     Print #lvIntFileNum, (lvContents) 'д�ļ�
     Close #lvIntFileNum '�ر��ļ�
     
     If notExistTable <> "" Then
        MsgBox "��" + notExistTable + "�����ڣ���"
     End If
     
     MsgBox "�ļ����ɳɹ����ļ�·��" + avFilePath
    
End Sub






