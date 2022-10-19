Attribute VB_Name = "Application"
'**************************************************************
' Application.bas - General API methods regarding the Application in general.
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

Option Explicit

''
' Retrieves the active window's hWnd for this app.
'
' @return Retrieves the active window's hWnd for this app. If this app is not in the foreground it returns 0.

Private Declare Function GetActiveWindow Lib "user32" () As Long

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    IsAppActive = (GetActiveWindow <> 0)

End Function

Public Sub LogError(ByVal Numero As Long, _
                    ByVal Descripcion As String, _
                    ByVal Componente As String, _
                    Optional ByVal Linea As Integer)

    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
    Dim File As Integer

    File = FreeFile
        
    Open App.path & "\Errores.log" For Append As #File
    
    Print #File, "Error: " & Numero
    Print #File, "Descripcion: " & Descripcion
        
    If LenB(Linea) <> 0 Then
        Print #File, "Linea: " & Linea

    End If
        
    Print #File, "Componente: " & Componente
    Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
    Print #File, vbNullString
        
    Close #File
    
    Debug.Print "Error: " & Numero & vbNewLine & "Descripcion: " & Descripcion & vbNewLine & "Componente: " & Componente & vbNewLine & "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
                
End Sub

