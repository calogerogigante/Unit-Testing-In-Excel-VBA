' ============================================================
'                         Assertions
'                 Auteur/Author : Calogero GIGANTE
' ============================================================
'                          ENGLISH
' ============================================================
' This is a simple Module (= standard module)
' Provides Assert objects + detailed messages (including both the value AND the VBA type).
' ============================================================
'                          FRANÇAIS
' ============================================================
' C'est un simple Module ( = Module standard)
' Fabriques d'objets Assert + messages détaillés (valeur ET type VBA).
' ============================================================
Option Explicit

'=============================================================
' Compares two values ​​using "=". Never throws an error: if the comparison is
' impossible (e.g., comparing a number to text, or comparing arrays), the result is a documented failure.
'-------------------------------------------------------------
' Compare deux valeurs avec "=". Ne lève jamais d'erreur : si la comparaison est
' impossible (ex. nombre contre texte, tableaux), le résultat est un échec expliqué.
Public Function AreEqual(ByVal expected As Variant, ByVal actual As Variant) As Assert
    Dim r As Assert
    Dim egaux As Boolean

    Set r = New Assert

    On Error GoTo Incomparables

    If IsNull(expected) Or IsNull(actual) Then
        egaux = IsNull(expected) And IsNull(actual)
    Else
        egaux = (expected = actual)
    End If

    r.AssertSuccessful = egaux
    If egaux Then
        r.AssertMessage = "OK"
    Else
        r.AssertMessage = "attendu " & Decrire(expected) & ", obtenu " & Decrire(actual) & Ecart(expected, actual)
    End If

    Set AreEqual = r
    Exit Function

Incomparables:
    r.AssertSuccessful = False
    r.AssertMessage = "valeurs incomparables : attendu " & Decrire(expected) & ", obtenu " & Decrire(actual)
    Set AreEqual = r
End Function

'=============================================================
' Comparison with a tolerance threshold, for floating-point numbers (Single, Double)
'-------------------------------------------------------------
' Comparaison avec tolérance, pour les nombres à virgule flottante (Single, Double)
Public Function AreNear(ByVal expected As Double, ByVal actual As Double, ByVal tolerance As Double) As Assert
    Dim r As Assert
    Dim ecart As Double

    Set r = New Assert
    ecart = Abs(expected - actual)
    r.AssertSuccessful = (ecart <= tolerance)

    If r.AssertSuccessful Then
        r.AssertMessage = "OK"
    Else
        r.AssertMessage = "attendu " & CStr(expected) & " +/- " & CStr(tolerance) & ", obtenu " & CStr(actual) & " (ecart = " & CStr(ecart) & ")"
    End If

    Set AreNear = r
End Function

'-------------------------------------------------------------
Public Function IsTrue(ByVal condition As Boolean, Optional ByVal message As String = "la condition est fausse") As Assert
    Dim r As Assert

    Set r = New Assert
    r.AssertSuccessful = condition
    If condition Then r.AssertMessage = "OK" Else r.AssertMessage = message

    Set IsTrue = r
End Function

'=============================================================
' "ERROR" result: to be used within a test's error handler when the
' code being tested crashes. The test runner will display this as [ERROR] (rather than [FAILURE]).
'-------------------------------------------------------------
' Résultat "ERREUR" : à utiliser dans le gestionnaire d'erreur d'un test, quand le
' code testé a planté. Le lanceur l'affichera comme [ERREUR] (et non [ECHEC]).
Public Function ErreurInattendue(ByVal numero As Long, ByVal description As String) As Assert
    Dim r As Assert

    Set r = New Assert
    r.AssertSuccessful = False
    r.AssertIsError = True
    r.AssertMessage = numero & " - " & description

    Set ErreurInattendue = r
End Function

' ------------------------------------------------------------
' Private display utilities
' ------------------------------------------------------------
' Outils privés d'affichage
' ------------------------------------------------------------

'=============================================================
' Value + VBA type, e.g.: 3 (Long), 'JD' (String), 2024-03-01 (Date)
'-------------------------------------------------------------
' Valeur + type VBA, ex. : 3 (Long), 'JD' (String), 2024-03-01 (Date)
Private Function Decrire(ByVal valeur As Variant) As String
    Dim texte As String

    On Error Resume Next

    Select Case VarType(valeur)
        Case vbNull
            texte = "Null"
        Case vbEmpty
            texte = "Empty"
        Case vbBoolean
            If valeur Then texte = "True" Else texte = "False"
        Case vbDate
            If valeur = Int(valeur) Then
                texte = Format$(valeur, "yyyy-mm-dd")
            Else
                texte = Format$(valeur, "yyyy-mm-dd hh:nn:ss")
            End If
        Case vbString
            texte = "'" & valeur & "'"
        Case Else
            texte = CStr(valeur)
    End Select

    Decrire = texte & " (" & TypeName(valeur) & ")"
End Function

'=============================================================
' For floating-point numbers: displays the difference (CStr rounds to 15 digits,
' meaning two distinct values ​​could otherwise appear identical)
'-------------------------------------------------------------
' Pour les nombres à virgule flottante : affiche l'écart (CStr arrondit à 15 chiffres,
' donc deux valeurs différentes peuvent s'afficher identiques sans cela)
Private Function Ecart(ByVal attendu As Variant, ByVal obtenu As Variant) As String
    On Error Resume Next

    If EstNumerique(attendu) And EstNumerique(obtenu) Then
        If VarType(attendu) = vbSingle Or VarType(attendu) = vbDouble Or _
           VarType(obtenu) = vbSingle Or VarType(obtenu) = vbDouble Then
            Ecart = " (ecart = " & CStr(Abs(CDbl(attendu) - CDbl(obtenu))) & ")"
        End If
    End If
End Function

'-------------------------------------------------------------
Private Function EstNumerique(ByVal valeur As Variant) As Boolean
    Select Case VarType(valeur)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            EstNumerique = True
    End Select
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
