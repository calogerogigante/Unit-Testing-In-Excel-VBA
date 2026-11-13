' ============================================================
'                         TestModele
'                 Auteur/Author : Calogero GIGANTE
' ============================================================
'                          ENGLISH
' ============================================================
' This is a standard module.
' BLANK TEMPLATE to be copied for each new test module.
' The module name MUST begin with "Test" so that the runner can find it.
' ============================================================
'                          FRANÇAIS
' ============================================================
' Ceci est un Module standard.
' MODELE VIERGE à copier pour chaque nouveau module de tests.
' Le nom du module DOIT commencer par "Test" pour que le runner (lanceur) le trouve.
' ============================================================
Option Explicit

' --- Hooks (optional: remove any you don't need) ---
' --- Hooks (facultatifs : supprime ceux dont vous n'avez pas besoin) ---

' Public Sub BeforeAll()
' End Sub
'
' Public Sub BeforeEach()
' End Sub
'
' Public Sub AfterEach()
' End Sub
'
' Public Sub AfterAll()
' End Sub

' --- Tool for testing an EXPECTED error: returns the error number, or 0 if none occurs ---
' --- Outil pour tester une erreur ATTENDUE : rend le numero d'erreur, ou 0 s'il n'y en a pas ---

' Private Function NumeroErreurDeMaFonction(ByVal x As Long) As Long
'     Dim r As Long
'
'     On Error GoTo Erreur
'     r = MaFonction(x)
'     NumeroErreurDeMaFonction = 0
'     Exit Function
'
' Erreur:
'     NumeroErreurDeMaFonction = Err.Number
' End Function


' --- A test: ALWAYS this format ---
' --- Un test : TOUJOURS cette forme ---
Public Function TestExempleADupliquer() As Assert
    On Error GoTo Erreur
    Set TestExempleADupliquer = AreEqual(4, 2 + 2)  ' = AreEqual( 4, FunctionToTest(2,2) )
    Exit Function

Erreur:
    Set TestExempleADupliquer = ErreurInattendue(Err.Number, Err.Description)
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
