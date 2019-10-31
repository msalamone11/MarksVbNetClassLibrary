Imports System.Runtime.CompilerServices

Namespace StringMethods

    Public Module StringMethods
        <Extension>
        Public Function StartsWithUpper(str As String) As Boolean
            If String.IsNullOrWhiteSpace(str) Then
                Return False
            End If

            Dim ch As Char = str(0)
            Return Char.IsUpper(ch)
        End Function
    End Module

End Namespace
