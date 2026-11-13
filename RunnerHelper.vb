' ============================================================
'                         RunnerHelper
'                 Auteur/Author : Calogero GIGANTE
' ============================================================
'                          ENGLISH
' ============================================================
' Executes a test WITHIN Excel and returns only a single string to PowerShell:
'   "OK|message", "KO|message" (failure), or "ERR|message" (error)
' ============================================================
'                          FRANÇAIS
' ============================================================
' Exécute un test DANS Excel et ne renvoie à PowerShell qu'une chaîne :
'   "OK|message"   "KO|message" (échec)   "ERR|message" (erreur)
' ============================================================
Option Explicit

Public Function ExecuterTest(ByVal nomMacro As String) As String
    Dim r As Assert

    On Error GoTo Erreur

    Set r = Application.Run(nomMacro)

    If r.AssertIsError Then
        ExecuterTest = "ERR|" & r.AssertMessage
    ElseIf r.AssertSuccessful Then
        ExecuterTest = "OK|" & r.AssertMessage
    Else
        ExecuterTest = "KO|" & r.AssertMessage
    End If
    Exit Function

Erreur:
    ExecuterTest = "ERR|" & Err.Number & " - " & Err.Description
End Function

'The MIT License (MIT)

'Copyright © 2026 <Calogero GIGANTE>

'Permission is hereby granted, free of charge, to any person
'obtaining a copy of this software and associated documentation
'files (the “Software”), to deal in the Software without restriction,
'including without limitation the rights to use, copy, modify,
'merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:

'The above copyright notice and this permission notice shall be
'included in all copies or substantial portions of the Software.

'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND,
'EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
'MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
'IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
'CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
'TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
'SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
