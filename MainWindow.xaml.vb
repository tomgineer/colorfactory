Imports System.Text.RegularExpressions
Class MainWindow
    Dim Darker(8) As TextBox
    Dim Lighter(8) As TextBox
    Dim RandomColors(8) As TextBox
    Dim UltraDarkGray As Color = ColorConverter.ConvertFromString("#141617")
    Dim LightGray As Color = ColorConverter.ConvertFromString("#E2E6E9")
    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        SetControls()
    End Sub

#Region "Colors"

    '   ultradark: #FF141617
    '   darkgray: #FF1C2023
    '   midgray: #FF141617
    '   gray: #FF282C2F
    '   lightergray: #FF627582
    '   lightgray: #FFE2E6E9 
    '   blue: #FF368BFC
    '   red: #FFDC3546

#End Region

    Private Function DarkenColor(ByVal c As Color, ByVal Perc As Integer) As Color
        Dim Factor As Single = 0
        Factor = 1 - Perc / 100
        Dim c2 As Color = Color.FromArgb(c.A, CInt((c.R * Factor)), CInt((c.G * Factor)), CInt((c.B * Factor)))
        Return c2
    End Function

    Private Function LightenColor(ByVal c As Color, ByVal Perc As Integer) As Color
        Dim Factor As Single = 0
        Factor = 1 + Perc / 100

        Dim NewR, NewG, NewB As Integer
        NewR = CInt((c.R * Factor))
        NewG = CInt((c.G * Factor))
        NewB = CInt((c.B * Factor))

        If NewR > 255 Then NewR = 255
        If NewG > 255 Then NewG = 255
        If NewB > 255 Then NewB = 255

        Dim c2 As Color = Color.FromArgb(c.A, NewR, NewG, NewB)
        Return c2
    End Function

    Private Sub SetControls()
        Darker(0) = Me.txtDarker10
        Darker(1) = Me.txtDarker20
        Darker(2) = Me.txtDarker30
        Darker(3) = Me.txtDarker40
        Darker(4) = Me.txtDarker50
        Darker(5) = Me.txtDarker60
        Darker(6) = Me.txtDarker70
        Darker(7) = Me.txtDarker80
        Darker(8) = Me.txtDarker90

        Lighter(0) = Me.txtLighter10
        Lighter(1) = Me.txtLighter20
        Lighter(2) = Me.txtLighter30
        Lighter(3) = Me.txtLighter40
        Lighter(4) = Me.txtLighter50
        Lighter(5) = Me.txtLighter60
        Lighter(6) = Me.txtLighter70
        Lighter(7) = Me.txtLighter80
        Lighter(8) = Me.txtLighter90

        RandomColors(0) = Me.txtColor1
        RandomColors(1) = Me.txtColor2
        RandomColors(2) = Me.txtColor3
        RandomColors(3) = Me.txtColor4
        RandomColors(4) = Me.txtColor5
        RandomColors(5) = Me.txtColor6
        RandomColors(6) = Me.txtColor7
        RandomColors(7) = Me.txtColor8
        RandomColors(8) = Me.txtColor9
    End Sub

    Private Function BrushFromHex(ByVal hexColorString As String) As SolidColorBrush
        Return CType((New BrushConverter().ConvertFrom(hexColorString)), SolidColorBrush)
    End Function

    Private Function Color2Hex(ByVal c As Color) As String
        Dim result As String
        result = c.ToString
        result = result.Remove(0, 3)
        Return "#" + result
    End Function

    Public Function Rand(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Static Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function

    Public Function getBrightness(ByVal c As Color) As Single
        '((Red value X 299) + (Green value X 587) + (Blue value X 114)) / 1000
        Dim Result As Single
        Result = ((CInt(c.R) * 299) + (CInt(c.G) * 587) + (CInt(c.B) * 114)) / 1000
        Return Math.Round(Result)

    End Function

    Private Sub btnPaste_Click(sender As Object, e As RoutedEventArgs)
        Dim Clip As String = My.Computer.Clipboard.GetText()
        Me.txtBase.Text = Clip

        If IsHex(Clip) Then
            DarkenLighten(Clip)
            Dim c As Color = ColorConverter.ConvertFromString(Clip)
            Me.txtBase.Background = New SolidColorBrush(c)
            If getBrightness(c) > 140 Then
                Me.txtBase.Foreground = New SolidColorBrush(UltraDarkGray)
            Else
                Me.txtBase.Foreground = New SolidColorBrush(LightGray)
            End If

        End If

    End Sub

    Private Function IsHex(inValue As String) As Boolean
        Dim regex As Regex = New Regex("^#[0-9A-F]{1,6}$")
        Dim match As Match = regex.Match(inValue)
        Return match.Success
    End Function

    Private Sub DarkenLighten(ByVal Hex As String)
        ' Darken Colors
        Dim c As Color = ColorConverter.ConvertFromString(Hex)
        Dim n As Integer = 0
        For Each Dark In Darker
            n += 1
            c = DarkenColor(c, n + 10)
            Dark.Text = Color2Hex(c)
            Dark.Background = New SolidColorBrush(c)
            If getBrightness(c) > 140 Then
                Dark.Foreground = New SolidColorBrush(UltraDarkGray)
            Else
                Dark.Foreground = New SolidColorBrush(LightGray)
            End If
        Next

        'Lighten Colors
        n = 0
        c = ColorConverter.ConvertFromString(Hex)
        For Each Light In Lighter
            n += 1
            c = LightenColor(c, n + 10)
            Light.Text = Color2Hex(c)
            Light.Background = New SolidColorBrush(c)
            If getBrightness(c) > 140 Then
                Light.Foreground = New SolidColorBrush(UltraDarkGray)
            Else
                Light.Foreground = New SolidColorBrush(LightGray)
            End If
        Next
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As RoutedEventArgs)
        Dim c As Color
        For Each RandomColor In RandomColors
            c = Color.FromArgb(255, Rand(0, 255), Rand(0, 255), Rand(0, 255))
            RandomColor.Text = Color2Hex(c)
            RandomColor.Background = New SolidColorBrush(c)

            If getBrightness(c) > 140 Then
                RandomColor.Foreground = New SolidColorBrush(UltraDarkGray)
            Else
                RandomColor.Foreground = New SolidColorBrush(LightGray)
            End If
        Next
    End Sub

End Class
