Option Strict On
Option Explicit On
Option Infer On

' import module
Namespace MyVariable
    Public Module ItemTextIndexes
        Public layoutProgressIdx As Integer = 0
        Public progressIdx As Integer = 0
    End Module

    Public Module printTextIndexes
        Public detectStatusIdx As Integer = 0
        Public outResultIdx As Integer = 0
        Public pbCheckResult As Integer = 0
    End Module

    Public Module ItemButtonIndexes
        Public instructionEditButtonIdx As Integer = 0
    End Module

    Public Module printButtonIndexes
        Public fileUpdateIdx As Integer = 0
        Public needCheckIdx As Integer = 0
    End Module
End Namespace