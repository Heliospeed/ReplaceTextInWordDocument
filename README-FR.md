# ReplaceTextInWordDocument

Ce projet C# montre comment utiliser la bibliothèque OpenXML SDK pour ouvrir un document Word (.docx) et remplacer du texte spécifique par un autre texte.

## Prérequis

- **.NET SDK** : Assurez-vous d'avoir le .NET SDK installé. Téléchargez-le depuis [le site officiel de .NET](https://dotnet.microsoft.com/download).
- **Open XML SDK** : Ce projet utilise la bibliothèque OpenXML SDK pour manipuler les documents Word.

### Installation de OpenXML SDK

Installez le package OpenXML SDK via NuGet :

```bash
dotnet add package DocumentFormat.OpenXml
```

## Structure du projet

Ce projet contient une classe `Docx` avec un constructeur qui prend en paramètre le chemin du fichier Word source.
Il y a ensuite une méthode `ReplaceText` qui prend en paramètre le texte à remplacer et le nouveau texte.
Et enfin une méthode `Save` qui prend en paramètre le chemin du fichier Word cible.

L'idée est pour le moment de recopier simplement cette classe dans un projet (et le faire évoluer en fonction de vos besoins).
J'ai volontairement fait un projet console pour faire une démo.

## Utilisation

### 1. Préparer le projet

Clonez ce dépôt et ouvrez-le dans votre éditeur C# préféré pour tester, ou recopier simplement le fichier `Docx.cs` dans votre projet.

### 2. Ajouter un document Word

Placez un document Word `.docx` dans le dossier du projet ou spécifiez un chemin vers un fichier Word existant. Dans le dossier `example`, j'ai créé un fichier Word contenant 2 remplacements a effectués. Le premier `#name` qui a été écrit de façon fragmenté et le second `#year` qui est présent à plusieurs emplacement dont une zone de texte.

### 3. Exécuter le code

Le code principal pour remplacer le texte est dans la méthode `ReplaceText`. Voici un exemple de code dans le fichier `Program.cs` permettant d'utiliser ce code :

```csharp
using ReplaceTextInWordDocument;

// Chargement du document Word
using var docx = new Docx("../example/example.docx");

// Remplacement consécutif dans un word (le rempalecement est volonairement insensible à la casse)
docx.ReplaceText("#name", "John Doe");
docx.ReplaceText("#year", DateTime.Now.Year.ToString());

// Sauvegarde de la nouvelle version du document
docx.Save("../example/exampleOut.docx");
```

### 4. Exécuter le programme

Dans votre terminal, naviguez jusqu'au dossier du projet et exécutez la commande suivante :

```bash
dotnet run
```

### 5. Vérifier le résultat

Le texte spécifié dans `votre_document.docx` aura été remplacé. Vous pouvez ouvrir le document Word pour vérifier le remplacement.

## Notes

- **Sensibilité à la casse** : Ce code est insensible à la casse. Si vous souhaitez le rendre sensible à la casse, adaptez la logique pour rechercher le texte en tenant compte de la casse.

## Ressources

- [Documentation de OpenXML SDK](https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk)
