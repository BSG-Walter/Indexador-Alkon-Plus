Attribute VB_Name = "Grh"
Option Explicit

Private Const GRH_DAT_FILE As String = "Graficos.ind"
Private Const OLD_FORMAT_HEADER As String = "Argentum Online by Noland-Studios."
Private Const OLD_FORMAT_INIT_FILE As String = "Inicio.con"

Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long
    
    Speed As Single
End Type

Private Type tCabecera 'Cabecera de los con
    desc As String * 255
    CRC As Long
    MagicWord As Long
End Type

Public Type tGameIni
    Puerto As Long
    Musica As Byte
    fX As Byte
    tip As Byte
    Password As String
    Name As String
    DirGraficos As String
    DirSonidos As String
    DirMusica As String
    DirMapas As String
    NumeroDeBMPs As Long
    NumeroMapas As Integer
End Type

Public GrhData() As GrhData

Public fileVersion As Long


Public Function LoadGrhData(ByVal path As String) As Boolean
On Error GoTo ErrHandler
    Dim Handle As Integer
    Dim MiCabecera As tCabecera
    
    'Set initial size
    ReDim GrhData(0) As GrhData
    
    Handle = FreeFile()
    
    If path = vbNullString Then Exit Function
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    If Not FileExists(path & GRH_DAT_FILE) Then
        MsgBox "The file " & path & GRH_DAT_FILE & " does not exist. A new one will be created with your work."
        Exit Function
    End If
    
    Open path & GRH_DAT_FILE For Binary Access Read Lock Write As Handle
    
    'Check file format! (The crappy header had to have some use after all!)
    Get Handle, , MiCabecera
    
    If Config.oldFormat Then
        LoadGrhData = LoadGrhDataOld(Handle, NumberOfGrhs(path))
        
        'No version available in old file format
        fileVersion = -1
    Else
        'We dont' have header, move back to the beginning
        Seek Handle, 1
        
        LoadGrhData = LoadGrhDataNew(Handle)
    End If
    
    Close Handle
Exit Function

ErrHandler:
    Close Handle
    
    MsgBox "An error occured while loading the grh data." & vbCrLf _
        & "Make sure file format is valid, and in case of using the old format, make sure the " _
        & OLD_FORMAT_INIT_FILE & " file is in the init folder"
End Function

''
' Old crappy format loading. Restricted to 2^15-1 grhs,
' stores animation speed in frames and other crappy stuff.
' Coded just for backwards compatibility, users should avoid using this format.
'
' @param    handle      Handle to the open file containing the grh data.
'                       The header should have allready been removed.
' @param    totalGrhs   The total number of grhs that could exist.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataOld(ByVal Handle As Integer, ByVal totalGrhs As Long) As Boolean
On Error GoTo ErrorHandler
    Dim grh As Integer
    Dim Frame As Long
    Dim tempint As Integer
    Dim max As Integer
    
    max = -1
    
    'Resize array
    ReDim GrhData(1 To totalGrhs) As GrhData
    
    'Open files
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint
    Get Handle, , tempint
    
    'Fill Grh List
    
    'Get first Grh Number
    Get Handle, , grh
    
    Do Until grh <= 0
        'Get highest grh number being used
        If grh > max Then
            max = grh
        End If
        
        With GrhData(grh)
            'Get number of frames
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            'Resize animation array
            ReDim .Frames(1 To .NumFrames) As Long
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                
                    Get Handle, , tempint
                    
                    'Old format uses integers
                    .Frames(Frame) = tempint
                    
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > totalGrhs Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get Handle, , tempint
                
                'Convert old speed to new one (time based)!
                .Speed = CSng(tempint) * .NumFrames * 1000 / 18
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get Handle, , tempint
                
                'Old format used ints, not longs.
                .FileNum = tempint
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get Handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                    
                Get Handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = grh
            End If
        End With
        
        'Get Next Grh Number
        Get Handle, , grh
    Loop
    
    Close Handle
    
    'Trim array
    ReDim Preserve GrhData(1 To max) As GrhData
    
    LoadGrhDataOld = True
Exit Function

ErrorHandler:
    LoadGrhDataOld = False
End Function

''
' Finds out the number of grhs for the old file format
'
' @param    path    The path to the folder in which the init file is stored.
'
' @return   The number of grhs that can exist at most.

Private Function NumberOfGrhs(ByVal path As String) As Long
    Dim N As Integer
    Dim GameIni As tGameIni
    Dim MiCabecera As tCabecera
    
    N = FreeFile
    
    Open path & OLD_FORMAT_INIT_FILE For Binary As #N
    
    Get N, , MiCabecera
    
    Get N, , GameIni
    
    Close N
    
    NumberOfGrhs = GameIni.NumeroDeBMPs
End Function

''
' Loads grh data using the new file format.
'
' @param    handle      Handle to the open file containing the grh data.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhDataNew(ByVal Handle As Integer) As Boolean
On Error GoTo ErrorHandler
    Dim grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    
    'Get file version
    Get Handle, , fileVersion
    
    'Get number of grhs
    Get Handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    While Not EOF(Handle)
        Get Handle, , grh
        
        With GrhData(grh)
            'Get number of frames
            Get Handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo ErrorHandler
            
            ReDim .Frames(1 To GrhData(grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get Handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo ErrorHandler
                    End If
                Next Frame
                
                Get Handle, , .Speed
                
                If .Speed <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo ErrorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo ErrorHandler
            Else
                'Read in normal GRH data
                Get Handle, , .FileNum
                If .FileNum <= 0 Then GoTo ErrorHandler
                
                Get Handle, , GrhData(grh).sX
                If .sX < 0 Then GoTo ErrorHandler
                
                Get Handle, , .sY
                If .sY < 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                Get Handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = grh
            End If
        End With
    Wend
    
    Close Handle
    
    LoadGrhDataNew = True
Exit Function

ErrorHandler:
    LoadGrhDataNew = False
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise (data can't be stored in the old file format, use new one).

Public Function SaveGrhDataOld(ByVal path As String) As Boolean
    Dim Handle
    Dim Frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    Handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As Handle
    
    MiCabecera.desc = OLD_FORMAT_HEADER
    
    'Write headers
    Put Handle, , MiCabecera
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    Put Handle, , tempint
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            'Index too big for this file format?
            If i > &H7FFF& Then
                Close Handle
                Kill path
                Exit Function
            End If
            
            Put Handle, , CInt(i)
            
            With GrhData(i)
                'Set number of frames
                Put Handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Put Handle, , CInt(.Frames(Frame))
                    Next Frame
                    
                    Put Handle, , CInt(.Speed * 0.018 / .NumFrames)
                Else
                    'Write in normal GRH data
                    Put Handle, , CInt(.FileNum)
                    
                    Put Handle, , .sX
                    
                    Put Handle, , .sY
                        
                    Put Handle, , .pixelWidth
                    
                    Put Handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close Handle
    
    SaveGrhDataOld = True
End Function

''
' Saves grh data using the old (and obsolete) file format. Shouldn't be used if possible.
' New format is valid with the new engine, included in Argentum Online 0.12.1
'
' @param    path    The complete path of the folde rin which to write the grh data file.
'                   If it existed it's deleted first.
'
' @return   True if the file was properly saved, False otherwise.

Public Function SaveGrhDataNew(ByVal path As String) As Boolean
    Dim Handle
    Dim Frame As Long
    Dim i As Long
    Dim tempint As Integer
    Dim MiCabecera As tCabecera
    
    'Make sure path is properly set
    If Right$(path, 1) <> "\" Then path = path & "\"
    
    path = path & GRH_DAT_FILE
    
    
    Handle = FreeFile()
    
    If FileExists(path) Then
        Call Kill(path)
    End If
    
    Open path For Binary Access Write As Handle
    
    'Increment file version
    fileVersion = fileVersion + 1
    
    Put Handle, , fileVersion
    
    Put Handle, , CLng(UBound(GrhData()))
    
    'Store Grh List
    For i = 1 To UBound(GrhData())
        If GrhData(i).NumFrames > 0 Then
            Put Handle, , i
            
            With GrhData(i)
                'Set number of frames
                Put Handle, , .NumFrames
                
                If .NumFrames > 1 Then
                    'Read a animation GRH set
                    For Frame = 1 To .NumFrames
                        Put Handle, , .Frames(Frame)
                    Next Frame
                    
                    Put Handle, , .Speed
                Else
                    'Write in normal GRH data
                    Put Handle, , .FileNum
                    
                    Put Handle, , .sX
                    
                    Put Handle, , .sY
                        
                    Put Handle, , .pixelWidth
                    
                    Put Handle, , .pixelHeight
                End If
            End With
        End If
    Next i
    
    Close Handle
    
    SaveGrhDataNew = True
End Function
