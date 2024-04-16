VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/  [Class] Directory Folder Object
'/  Created By:     Ryan Tamashiro
'/  Contact Info:   Tamashiroryan@gmail.com
'/**********************************************************************************
'/  **README**
'/
'/  Required Library References
'/  1. Microsoft Scripting Runtime
'/
'/
'/  Class Properties:
'/  1.  Path                    / Working Directory Path
'/  2.  FileCount               / Number of Files in Directory
'/  3.  FolderCount             / Number of Folders in Directory
'/  4.  FilterCount             / Number of Files of Specified Type in Directory
'/  5.  Size                    / Size of Directory in MB
'/
'/
'/  Class Methods:
'/  1.  PrintFiles              / Print All Files to Immediate Window
'/  2.  PrintFolders            / Print All Folders to Immediate Window
'/  3.  FilterFiles             / Filters Specified File Types Into New Collection
'/  4.  LastModified            / Returns Name of Last Modified File
'/**********************************************************************************
Option Explicit

Private pvtPath                 As String
Private pvtFiles                As New Collection
Private pvtFldrs                As New Collection
Private pvtFiltered             As Collection
Private pvtFileCt               As Long
Private pvtFldrCt               As Long
Private pvtDirSize              As Double

'########################### [Class|Event] <Initialize> ############################
Private Sub Class_Initialize():
'/  Created On: 03/20/2020                  Last Modified:  03/20/2020
'/=================================================================================
'/  Description:
'/
'/      Init Class Variables:
'/      [Set] pvtFiles
'/      [Set] pvtFldrs
'/      [Set] pvtFileCt
'/      [Set] pvtFldrCt
'/      [Set] pvtDirSize
'/
'/=================================================================================
Const MsgInit   As String = "[Init] Directory Class Object"
Dim bnPathSet   As Boolean

    Call SetDirPath(Path:=CurDir())
    bnPathSet = Len(pvtPath) > 0
    If bnPathSet Then
        Call ListDirectory
    End If
End Sub
'###################################################################################


'/**********************************************************************************
'/                          [Class] Properties                                     *
'/**********************************************************************************

Public Property Let Path(Path As String):
    Call SetDirPath(Path:=Path)
    Debug.Print Now(); "[Set|Property] Change Dir Path: <"; pvtPath; ">"
    Call ResetDirectory
    Call ListDirectory
End Property

Public Property Get Path() As String
    Path = pvtPath
End Property

Public Property Get Files() As Collection
    Files = pvtFiles
End Property

Public Property Get Folders() As Collection
    Folders = pvtFldrs
End Property

Public Property Get FileCount() As Long
    FileCount = pvtFileCt
End Property

Public Property Get FolderCount() As Long
    FolderCount = pvtFldrCt
End Property


'/**********************************************************************************
'/                          [Class] Methods                                        *
'/**********************************************************************************
Private Sub SetDirPath(Path As String)
'/  Created On:     03/20/2020                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
Const MsgPathErr    As String = "[Error] Invalid/Empty Directoy Path <"
Const MsgPathSet    As String = "[Set|Property] Path: <"
Dim bnValidPath     As Boolean
Dim strDirPath      As String

    strDirPath = PathRepair(Path:=Path)
    bnValidPath = ValidatePath(Path:=strDirPath)
    If bnValidPath Then
        Debug.Print Now(); MsgPathSet; strDirPath; ">"
        pvtPath = strDirPath
        Exit Sub
    End If
Debug.Print Now(); MsgPathErr; Path; ">"
End Sub


Private Sub ListDirectory():
'/  Created On:     06/03/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
Const MsgFiles  As String = "[Set|Property] Files - Count: "
Const MsgFldrs  As String = "[Set|Property] Fldrs - Count: "
Dim bnIsFldr    As Boolean
Dim strDirRtn   As String
Dim strFullPath As String

    strDirRtn = Dir(pvtPath, vbDirectory)
    Do While (strDirRtn <> vbNullString)
        Select Case strDirRtn
            Case Is = ".", ".."
            Case Else
                strFullPath = pvtPath & strDirRtn
                bnIsFldr = (GetAttr(strFullPath) And vbDirectory) > 0
                If bnIsFldr Then
                    pvtFldrs.Add Item:=strDirRtn
                Else
                    pvtFiles.Add Item:=strDirRtn
                End If
        End Select
        strDirRtn = Dir()
    Loop
    pvtFileCt = pvtFiles.Count
    pvtFldrCt = pvtFldrs.Count
End Sub


Private Sub ResetDirectory():
'/  Created On:     03/20/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
    On Error Resume Next
    pvtFileCt = 0
    pvtFldrCt = 0
    Set pvtFiles = Nothing
    Set pvtFldrs = Nothing
    On Error GoTo 0
End Sub


Public Sub PrintFiles():
'/  Created On:     06/03/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
Const MsgPrint  As String = "[Print] Directory Files"
Const MsgStats  As String = "<Results> File Count: "
Dim bnPathSet   As Boolean
Dim vFile       As Variant
Dim i           As Long

    Debug.Print Now(); MsgPrint
    bnPathSet = Len(pvtPath) > 0
    If bnPathSet Then
         i = 0
         For Each vFile In pvtFiles
            i = i + 1
            Debug.Print vbTab; i; "- <File> "; vFile
        Next vFile
    End If
Debug.Print vbTab; MsgStats; pvtFileCt; vbCrLf
End Sub


Public Sub PrintFolders():
'/  Created On:     06/03/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
Const MsgPrint  As String = "[Print] Directory Files"
Const MsgStats  As String = "<Results> Fldr Count: "
Dim bnPathSet   As Boolean
Dim vFldr       As Variant
Dim i           As Long

    Debug.Print Now(); MsgPrint
    bnPathSet = Len(pvtPath) > 0
    If bnPathSet Then
        i = 0
        For Each vFldr In pvtFldrs
            i = i + 1
            Debug.Print vbTab; i; "- <Fldr> "; vFldr
        Next vFldr
    End If
Debug.Print vbTab; MsgStats; pvtFldrCt; vbCrLf
End Sub


Private Function PathRepair(Path) As String
'/  Created On:     06/03/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
    PathRepair = Trim(Path)
    If Right(Trim(PathRepair), 1) <> "\" Then
        PathRepair = Trim(Path) & "\"
        Exit Function
    End If
End Function


Public Function ValidatePath(Path As String) As Boolean
'/  Created On:     06/03/2018                 Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
'/  Description:    Validates Arg Path Agains Below Criteria:
'/                      1. Path Must Not be Null (Length Zero)
'/                      2. Path Format Ends w/ Back Slash
'/                      3. Successful Connection (Path Call Rtn Not Null)
'/                      **[Validate] Conditon: All Criterial Are True
'/----------------------------------------------------------------------------------
Dim bnPathNotNull   As Boolean
Dim bnValidFormat   As Boolean
Dim bnSuccConn      As Boolean
    
    bnPathNotNull = (Path <> vbNullString)
    bnValidFormat = (Right(Trim(Path), 1) = "\")
    bnSuccConn = (Dir(Path, vbNormal) <> vbNullString)
    ValidatePath = (bnPathNotNull And bnValidFormat And bnSuccConn)
End Function


Public Sub ConsoleLog(arg As Variant):
'/  Created On: 03/20/2020                  Last Modified:  03/20/2020
'/----------------------------------------------------------------------------------
'/  Description:    Write to Console Log Only When In Debugging Mode
'/
'/
'/----------------------------------------------------------------------------------
#If Debugging = 0 Then



#End If
End Sub
