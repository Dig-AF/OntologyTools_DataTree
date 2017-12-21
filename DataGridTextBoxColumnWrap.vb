Public Class DataGridTextBoxColumnWrap
    Inherits DataGridTextBoxColumn

    Protected Overloads Overrides Sub Paint(ByVal g As System.Drawing.Graphics, _
                                                    ByVal bounds As System.Drawing.Rectangle, _
                                                    ByVal source As System.Windows.Forms.CurrencyManager, _
                                                    ByVal rowNum As Integer, _
                                                    ByVal backBrush As System.Drawing.Brush, _
                                                    ByVal foreBrush As System.Drawing.Brush, _
                                                    ByVal alignToRight As Boolean)

        Dim stringSize As SizeF
        Dim stringFont As Font
        Dim columnWidth As Integer
        Dim stringValue As String
        Dim stringDisplay As String
        Dim stringYCoord As Single
        Dim stringHeight As Single
        Dim stringArray As ArrayList
        Dim intStringLoop As Integer

        g.FillRectangle(backBrush, bounds) 'Paint the rectangle for the cell using the backBrush 
        stringFont = Me.DataGridTableStyle.DataGrid.Font 'Get the font used by the column(style) 
        columnWidth = bounds.Width 'Get the width of the column 

        stringSize = g.MeasureString("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890", stringFont)
        stringHeight = stringSize.Height
        stringValue = GetText(source, rowNum) 'Get the text to be displayed in the cell 
        stringArray = GetArrayList(g, stringValue, stringFont, columnWidth)

        For intStringLoop = 0 To stringArray.Count - 1
            stringDisplay = CType(stringArray.Item(intStringLoop), String)
            stringYCoord = bounds.Y + (intStringLoop * stringHeight)
            g.DrawString(stringDisplay, stringFont, foreBrush, bounds.X, stringYCoord)
        Next

    End Sub

    Public Sub New()
        ' Initialise the object and set its default properities 
        MyBase.New()
        Me.NullText = ""
        Me.ReadOnly = False

    End Sub

    'Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal readOnly1 As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)
    '    ' This particular column will always be readonly and non-editable - therefore do not pass the event over 
    '    Return
    'End Sub

#Region "Helper Methods"

    Function GetText(ByVal source As System.Windows.Forms.CurrencyManager, _
                              ByVal rowNum As Integer) As String

        Dim objValue As Object = GetColumnValueAtRow(source, rowNum)
        If objValue Is System.DBNull.Value Then
            Return Me.NullText
        Else
            Return CType(objValue, String)
        End If

    End Function

    Function GetArrayList(ByVal g As Graphics, ByVal stringValue As String, _
                                    ByVal stringFont As Font, ByVal columnWidth As Integer) As ArrayList
        Dim intCurrentChar As Integer
        Dim stringCurrentChar As String
        Dim stringDisplay As String = ""
        Dim intSpace As Integer
        Dim stringSize As SizeF
        Dim stringArray As New ArrayList

        Do Until stringValue = ""
            For intCurrentChar = 0 To stringValue.Length - 1
                stringCurrentChar = stringValue.Substring(intCurrentChar, 1)
                stringSize = g.MeasureString(stringDisplay & stringCurrentChar, stringFont)
                If stringSize.Width <= columnWidth Then
                    stringDisplay &= stringCurrentChar
                    If intCurrentChar = stringValue.Length - 1 Then
                        stringArray.Add(stringDisplay)
                        stringDisplay = ""
                        stringValue = ""
                        Exit For
                    End If
                Else
                    'Figure out where the first space preceding this location is... 
                    intSpace = stringDisplay.LastIndexOf(" ")
                    If intSpace >= 0 Then
                        'we found the last space 
                        intCurrentChar = intSpace + 1
                        stringDisplay = stringDisplay.Substring(0, intSpace + 1)
                        stringArray.Add(stringDisplay)
                        stringDisplay = ""
                        stringValue = stringValue.Substring(intCurrentChar)
                        Exit For
                    Else

                        stringArray.Add(stringDisplay)
                        stringDisplay = ""
                        stringValue = stringValue.Substring(intCurrentChar)
                        Exit For
                    End If
                End If
            Next
        Loop

        Return stringArray

    End Function

#End Region

End Class