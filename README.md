# UpdateError

## Description
UpdateError est une application Windows Forms développée en VB.NET qui permet de vérifier et corriger les numéros de traçabilité dans une base de données SQL Server. L'application est spécifiquement conçue pour valider que les numéros de traçabilité respectent les règles suivantes :
- 9 caractères si la quantité est égale à 1
- 16 caractères si la quantité est supérieure à 1

## Fonctionnalités
- Connexion sécurisée à une base de données SQL Server (authentification Windows ou SQL)
- Configuration de la connexion via une interface graphique dédiée
- Stockage crypté des informations de connexion
- Recherche de commandes par numéro de commande et numéro de ligne de mission
- Visualisation des lignes en erreur (numéros de traçabilité non conformes)
- Édition directe des numéros de traçabilité
- Interface utilisateur intuitive aux couleurs d'Eurodislog

## Prérequis
- Windows 10 ou supérieur
- .NET 6.0 Runtime
- Accès à une base de données SQL Server contenant les tables nécessaires

## Installation
1. Téléchargez la dernière version depuis le système de gestion de versions
2. Décompressez l'archive dans le dossier de votre choix
3. Lancez l'application en exécutant `UpdateError.exe`

## Configuration initiale
Lors du premier démarrage, l'application vous demandera de configurer la connexion à la base de données :
1. Connectez-vous en tant qu'administrateur (identifiants par défaut : admin/admin123)
2. Renseignez les informations de connexion à votre serveur SQL
3. Testez la connexion avant d'enregistrer

## Utilisation
1. **Recherche de commandes** :
   - Saisissez un numéro de commande dans le champ dédié
   - Vous pouvez également spécifier un numéro de ligne de mission (optionnel)
   - Appuyez sur le bouton "Rechercher" ou sur la touche "Entrée"

2. **Consultation des résultats** :
   - Les lignes en erreur sont affichées en rouge
   - Passez la souris sur les cellules en erreur pour voir le message détaillé

3. **Modification des numéros de traçabilité** :
   - Double-cliquez sur la cellule du numéro de traçabilité à modifier
   - Ou sélectionnez une ligne et utilisez le menu contextuel (clic droit)
   - Dans la fenêtre d'édition, corrigez le numéro selon les règles affichées
   - Confirmez avec le bouton "Enregistrer"

## Structure du projet
- **Program.vb** : Point d'entrée de l'application
- **DatabaseConnection.vb** : Gestion de la connexion à la base de données
- **DataAccess.vb** : Méthodes d'accès aux données (requêtes SQL)
- **Cryptage.vb** : Fonctions de cryptage/décryptage pour les informations sensibles
- **FormMain.vb** : Interface principale de l'application
- **FormLogin.vb** : Authentification administrateur
- **FormConnexion.vb** : Configuration de la connexion à la base de données
- **FormEditionExoTrak.vb** : Édition des numéros de traçabilité

## Dépannage
- **Erreur de connexion à la base de données** : Vérifiez les paramètres de connexion et assurez-vous que le serveur SQL est accessible
- **Numéro de traçabilité non modifiable** : Assurez-vous d'avoir les droits nécessaires sur la table SPE_EXO
- **Erreur d'affichage** : Redimensionnez la fenêtre ou relancez l'application

## Sécurité
- Les informations de connexion sont stockées de manière cryptée
- L'authentification administrateur est requise pour modifier la configuration
- Les requêtes SQL utilisent des paramètres pour prévenir les injections SQL

## Maintenance
Pour modifier les règles de validation des numéros de traçabilité, consultez la méthode `EstExoTrakValide` dans le fichier `DataAccess.vb`.

## Contact
Pour toute question ou signalement de bug, veuillez contacter le service informatique d'Eurodislog.

---
© 2025 Eurodislog - Tous droits réservés