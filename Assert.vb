' ============================================================
'                           Assert
'                 Auteur/Author : Calogero GIGANTE
' ============================================================
'                          ENGLISH
' ============================================================
' This is a Class module.
' Make in Visual Basic Editor : Insert > Class Module, then
' rename to "Assert" using F4.
'
' Test result. Three possible outcomes :
'   - Success: AssertSuccessful = True
'   - Failure: AssertSuccessful = False, AssertIsError = False
'   - Error:   AssertIsError = True (the code being tested crashed)
' ============================================================
'                          FRANÇAIS
' ============================================================
' C'est un Module de classe.
' Dans le Visual Basic Editor, faire : Insertion > Module de
' classe, et renommer en "Assert" via F4.
'
' Résultat d'un test. Trois issues possibles :
'   - réussite : AssertSuccessful = True
'   - échec    : AssertSuccessful = False, AssertIsError = False
'   - erreur   : AssertIsError = True (le code testé a planté)
' ============================================================
Option Explicit

Public AssertSuccessful As Boolean
Public AssertMessage As String
Public AssertIsError As Boolean

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
