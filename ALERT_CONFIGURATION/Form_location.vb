Public Class Form_location

    Public Shared x_position As Integer = 450
    Public Shared y_position As Integer = 250

    ''TESTE PARA REDIMENSIONAR ECRÂ
    '###################################TESTE###########################################################
    ''Bloco para redimensionar ecrã
    'Dim DesignScreenWidth As Integer = 1600
    'Dim DesignScreenHeight As Integer = 900

    'Dim CurrentScreenWidth As Integer
    'Dim CurrentScreenHeight As Integer

    'Dim RatioX As Double
    'Dim RatioY As Double

    '+9 é o offset porque o maximize coloca no ponto -8, -8
    'If (Me.Location.X + 9) > DesignScreenWidth Then

    '        'Ecrã secundário 
    '        CurrentScreenWidth = Screen.AllScreens(1).Bounds.Width
    '        CurrentScreenHeight = Screen.AllScreens(1).Bounds.Height

    '    Else

    '        'Ecrã primário
    '        CurrentScreenWidth = Screen.AllScreens(0).Bounds.Width
    '        CurrentScreenHeight = Screen.AllScreens(0).Bounds.Height

    '    End If

    'If CurrentScreenWidth < DesignScreenWidth Then

    '        RatioX = CurrentScreenWidth / DesignScreenWidth
    '        RatioY = CurrentScreenHeight / DesignScreenHeight

    '        MsgBox(RatioX)
    '        MsgBox(RatioY)

    '        For Each iControl In Me.Controls
    'With iControl
    'If (.GetType.GetProperty("Width").CanRead) Then .Width = CInt(.Width * RatioX)
    '                If (.GetType.GetProperty("Height").CanRead) Then .Height = CInt(.Height * RatioY)
    '                If (.GetType.GetProperty("Top").CanRead) Then .Top = CInt(.Top * RatioX)
    '                If (.GetType.GetProperty("Left").CanRead) Then .Left = CInt(.Left * RatioY)
    '            End With
    'Next

    'End If

    'Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
    '###################################################################################################

End Class
