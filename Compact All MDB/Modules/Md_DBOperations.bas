Attribute VB_Name = "Md_DBOperations"
'
'___________________________________________________________________________
' Program name      : KhompAllMDB.
' Description       : An easy way to gain some space on your HD,
'                     by compacting Access Databases.
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.01.19
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.01.19
'___________________________________________________________________________
' TODO :
'       - .
'       -
'___________________________________________________________________________
'
' Module description:
' Basic Micorsoft DAO database operations.
'___________________________________________________________________________
'


Function subRepairDataBase(strDataBasePath As String) As Integer
    Dim strDBName As String
    
    On Error GoTo Err_Handler
    subRepairDataBase = 1           'Everything is ok default.
    strDBName = strDataBasePath
    RepairDatabase strDBName
    Exit Function
Err_Handler:
    MsgBox "An error occured while repairing:" + CStr(strDataBasePath) + vbCrLf + "Description:" + Err.Description, vbCritical + vbOKOnly, "Repair error..."
    subRepairDataBase = 0           'Error occured.
End Function

Function funcCompactDataBase(strDataBasePath As String) As Double
    Dim strDBName As String, strDBNameBack As String
    Dim dblCompactedDBSize As Double
    Dim a
    
    On Error GoTo Err_Handler
    strDBName = strDataBasePath
    a = FileLen(strDBName)
    'First we create a backup name for the original MDB file.
    strDBNameBack = Left(strDataBasePath, Len(strDataBasePath) - 4) + "Back" + Right(strDataBasePath, 4)
    
    'We try to compact this file
    'CompactDatabase srcName,dstName
    CompactDatabase strDBName, strDBNameBack
    'strDBNameBack should be the compacted version of strDBName
    
    'If the first compact produces a strDBNameBack file with a
    'length>0 we assume we compacted properly.
    If FileLen(strDBNameBack) > 0 Then
        Kill strDBName                              'We kill the original file.
        CompactDatabase strDBNameBack, strDBName    'We recompact from the backup to the original file.
        If FileLen(strDBName) > 0 Then              'If size>0 then we assume it was successfull.
            Kill strDBNameBack                      'Kill the bakcup file.
            dblCompactedDBSize = FileLen(strDBName) 'Return the size of the compacted file.
            funcCompactDataBase = dblCompactedDBSize
        Else
            FileCopy strDBNameBack, strDBName       'Else we copy the backup to the original
            Kill strDBNameBack                      'and kill the backup file.
            funcCompactDataBase = -1                'Return -1 if failed.
        End If
    Else
        Kill strDBNameBack                          'First compact failed. We kill the bakcup file.
        funcCompactDataBase = -1                    'Return -1: failed
    End If
    Exit Function
Err_Handler:
    MsgBox "An error occured while compacting:" + CStr(strDataBasePath) + vbCrLf + "Description:" + Err.Description, vbCritical + vbOKOnly, "Compact error..."
End Function
