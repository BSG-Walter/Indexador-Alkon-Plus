Attribute VB_Name = "Desindexado"
'funciones robadas y adaptadas del DIU

Public Function Desindexar0120(ByVal grh As Long) As String
On Error GoTo ErrorHandler
    Dim Datos As String
        With GrhData(grh)
            If .NumFrames > 1 Then
                Datos$ = CStr(.NumFrames)
            
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Datos$ = Datos$ & "-" & CStr(.Frames(Frame))
                Next Frame
                
                'If .Speed <= 0 Then GoTo ErrorHandler

                Datos$ = Datos$ & "-" & CStr(.Speed)
            Else
                'Read in normal GRH data
                'If .FileNum <= 0 Then GoTo ErrorHandler
                
                'If .sX < 0 Then GoTo ErrorHandler
                
                'If .sY < 0 Then GoTo ErrorHandler
                
                'If .pixelWidth <= 0 Then GoTo ErrorHandler
                
                'If .pixelHeight <= 0 Then GoTo ErrorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
    
                .Frames(1) = grh
                Datos$ = "1-" & CStr(.FileNum) & "-" & CStr(.sX) & "-" & CStr(.sY) & "-" & CStr(.pixelWidth) & "-" & CStr(.pixelHeight)
            End If
        End With
        If LenB(Datos$) <> 0 Then
            Desindexar0120 = "Grh" & CStr(grh) & "=" & Datos$ & vbCrLf
        Else
            Desindexar0120 = "Error"
        End If
    
Exit Function

ErrorHandler:

End Function
