Attribute VB_Name = "EventosAoyind"
Option Explicit

Public Const MAX_EVENT_SIMULTANEO As Byte = 5

Public Enum eModalityEvent

    CastleMode = 1
    DagaRusa = 2
    DeathMatch = 3
    Enfrentamientos = 4

End Enum

Public Function strModality(ByVal Modality As eModalityEvent) As String

    Select Case Modality

        Case eModalityEvent.CastleMode
            strModality = "CastleMode"
            
        Case eModalityEvent.DagaRusa
            strModality = "DagaRusa"
            
        Case eModalityEvent.DeathMatch
            strModality = "DeathMatch"
            
        Case eModalityEvent.Enfrentamientos
            strModality = "Duelos"

    End Select

End Function

Public Function ModalityByte(ByVal Modality As String) As String

    Select Case Modality

        Case "CASTLEMODE"
            ModalityByte = 1
            
        Case "DAGARUSA"
            ModalityByte = 2
            
        Case "DEATHMATCH"
            ModalityByte = 3
            
        Case "1VS1", "2VS2", "3VS3", "4VS4", "5VS5", "6VS6", "7VS7", "8VS8", "9VS9", "10VS10", "11VS11", "12VS12", "13VS13", "14VS14", "15VS15", "20VS20", "25VS25"
              
            ModalityByte = 4

    End Select

End Function

