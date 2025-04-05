# Commande Vocale pour Ouvrir des Fichiers Excel Automatiquement 🎤📂

Ce document présente une méthode pour configurer une commande vocale qui ouvre automatiquement des fichiers Excel et active leur modification.

---

## Prérequis 🛠️

- **Système d'exploitation :** Windows 10 ou supérieur.
- **Reconnaissance vocale Windows :** Activée via le module de Reconnaissance vocale.
- **Scripts et macros :** Connaissances de base en VBA ou l'utilisation de fichiers BAT.

---

## Méthode 1 : Utiliser la Reconnaissance Vocale Windows avec un Script BAT 📜

### Étapes :

1. **Création du Script VBA**  
   Créez un fichier VBScript (par exemple, `buenos_dias.vbs`) qui ouvre vos fichiers Excel et active les modifications.  
   Exemple de code :
   ```vbscript
   Dim excelApp
   Set excelApp = CreateObject("Excel.Application")
   excelApp.Visible = True

   ' Ouvrir le premier fichier Excel
   Dim wb1
   Set wb1 = excelApp.Workbooks.Open("C:\chemin\vers\fichier1.xlsx")
   On Error Resume Next
   wb1.Unprotect "votre_mot_de_passe"
   On Error GoTo 0

   ' Ouvrir le second fichier Excel
   Dim wb2
   Set wb2 = excelApp.Workbooks.Open("C:\chemin\vers\fichier2.xlsx")
   On Error Resume Next
   wb2.Unprotect "votre_mot_de_passe"
   On Error GoTo 0
   ```

2. **Création du Fichier BAT**  
   Créez un fichier BAT (par exemple, `buenos_dias.bat`) qui appelle le script VBScript :
   ```bat
   @echo off
   cscript //nologo "C:\chemin\vers\buenos_dias.vbs"
   ```

3. **Configuration de la Commande Vocale**  
   Utilisez le module de **Reconnaissance vocale Windows** ou **Windows Speech Recognition Macros** pour créer une commande personnalisée (ex. : "Bonjour Excel") qui exécute le fichier BAT créé.

---

## Méthode 2 : Utiliser des Outils Tiers (comme Dragon NaturallySpeaking) 🐉

### Étapes :

1. **Configuration de l'Outil de Reconnaissance Vocale**  
   Installez et configurez Dragon NaturallySpeaking (ou un outil similaire) pour créer un profil de commandes vocales.

2. **Assignation de la Commande Vocale**  
   Définissez une commande (ex. : "Bonjour Excel") qui exécute le fichier BAT ou une macro intégrée.

---

## Remarques Importantes ⚠️

- **Chemins et Mots de Passe :**  
  Vérifiez que les chemins vers les fichiers Excel et les mots de passe utilisés sont corrects.

- **Sécurité :**  
  Assurez-vous que votre système autorise l'exécution de scripts et que les paramètres de sécurité sont bien configurés.

- **Tests :**  
  Testez votre configuration pour vous assurer que la commande vocale déclenche bien l'ouverture des fichiers et l'activation de la modification.

---

## Conclusion 🚀

En suivant ces méthodes, vous pouvez automatiser l'ouverture de fichiers Excel et activer leur modification via une commande vocale, facilitant ainsi vos tâches quotidiennes.
