# Unit-Testing-In-Excel-VBA

ENGLISH : ============================================
This framework enables the implementation of unit tests for VBA in Excel.

If the "business" functions (the ones to be tested) are located in a standard module named "Metier" then the corresponding test functions must be placed in a standard module named "TestMetier."

"Assert" must be a class module, while "Assertions" and "RunnerHelper" should be standard modules.

Tests are executed via a PowerShell script located in the same folder as the .xlsm file containing the business logic.

This framework was inspired by the following article:
https://dev.to/michaelneu/unit-testing-vba-macros-npb

The author also consulted Claude-AI on several occasions for guidance to ensure the framework was well-structured.
The code is distributed for free under the MIT License.

FRANÇAIS : -===========================================
Ce framework permet de mettre en place des tests unitaires en VBA dans Excel.

Si les fonctions "métier" (à tester) se trouvent dans le module standard "Metier", les fonctions de tests doivent se trouver dans le module standard "TestMetier".

"Assert" doit être un module de classe. "Assertions" et "RunnerHelper" sont des modules standards.

Les tests se lancent depuis un script powershell situé dans le même dossier où se trouve le fichier .xlsm contenant le code métier.

Ce framework a été inspiré de cet article :
https://dev.to/michaelneu/unit-testing-vba-macros-npb

Et l'auteur a consulté quelques fois Claude-IA et ses conseils pour s'assurer de la bonne tenue du framework.
Le code est distribué gratuitement sous la MIT-licence.

