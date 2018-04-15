 Public screenSize As Size

    Private Function ColorInBitmap(ByVal image As Bitmap, ByVal Target As Color) As Boolean
        Dim flag As Boolean = False
        Dim width As Integer = image.Width
        Dim height As Integer = image.Height
      

        For h As Integer = 0 To height - 1
            For w As Integer = 0 To width - 1
                If (w > width) Then
                    flag = False
                    w = 0
                    If (h > height) Then
                        flag = False
                        Exit For
                    ElseIf (image.GetPixel(w, h).Equals(Target)) Then
                        ' MsgBox("fond in Y")
                        flag = True
                    End If
                ElseIf (image.GetPixel(w, h).Equals(Target)) Then
                    '  MsgBox("fond in X")
                    flag = True
                End If
            Next
        Next
        Return flag

    End Function


    Public Function CheckColor(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)

        Using bitmap As System.Drawing.Bitmap = New System.Drawing.Bitmap(width, height)
            Using graphic As Graphics = Graphics.FromImage(bitmap)
                Dim point As System.Drawing.Point = New System.Drawing.Point(x, y)
                Dim point1 As System.Drawing.Point = New System.Drawing.Point(x + width, y + height)
                graphic.CopyFromScreen(point, point1, screenSize)
                Dim bitmap1 As System.Drawing.Bitmap = New System.Drawing.Bitmap(width, height)
                Dim graphic1 As Graphics = Graphics.FromImage(bitmap1)
                point1 = New System.Drawing.Point(x, y)
                graphic1.CopyFromScreen(point1, System.Drawing.Point.Empty, bitmap1.Size)
                If (ColorInBitmap(bitmap1, Color.FromArgb(r, g, b))) Then
                    Return True
                Else
                    Return False
                End If
            End Using
        End Using
    End Function
    
    
    
    
'USAGE
'================
'If CheckColor(Val(ColorRedValue.Value), Val(ColorGreenValue.Value), Val(ColorBlueValue.Value), Val(x1.Text), Val(y1.Text), Val(width1.Text), Val(height1.Text)) Then
           '''
'End If





