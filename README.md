Aride Online Creator
====================

Ceci est une verison de Frog Creator adaptée pour le développement du MMORPG Aride Online.
Vous pouvez utiliser tout ou partie de ce projet pour vos projets respectifs.

Version de Frog Creator utilisée : V0.6.2

Technologies utilisées :
* VB6
* DirectX7
* Java
* Python
* MySQL
* PHPBB

Voici une liste non exausthive des fonctionnalités qui se trouvent dans ce code.	
Capacités techniques :
----------------------
* Chargement de pointeur de souris dynamique
* Possibilité d'utiliser une icône client haute résolution avec tailles multiples
* Déserialisation de fichiers binaires VB6 en objets Java
* Communication VB6 <-> Java via envoi de données binaires
* Hooking des événements souris permettant, par exemple :
	* Drag and drop
	* Prise en compte de la molette
* Chargement dynamique des textures (Economie de mémoire ram)

Mécanismes :
----------------------
* Cartes à taille variable
* Météo :
	* Aléatoire et fonction des préférences indiquées dans l'éditeur de map
	* Fonctionnement par zone pour être au plus proche d'un comportement réel

* Déplacement des Npcs via un algorithme A*
* Mécanisme d'adoption des Npcs pour qu'ils deviennent des familiers
* Possibilité d'avoir plusieurs instances d'une même carte avec associations des instances pour créer des instances de donjon
* Modifications sur les types d'armement
	* Les armes de lancer se consomment elles-mêmes (ex : caillou)
	* Les armes de tir consomment des munitions dans l'inventaire (ex : arc)
* Mécanisme de fatigue avec baisse de la vision

Calculs :
----------------------
* Calculs et évolutions de l'environnement côté Serveur
* Le client prédit faiblement les évolutions et le serveur corrige si besoin

Cryptage :
----------------------
* Mots de passe :
	* Vous devez modifier la méthode de cryptage de votre forum PHPBB et adopter la même méthode du côté du client.
	* Vous comprendrez que notre méthode de cryptage a volontairement été supprimée de la version distribuée.
		
* Images :
	* Cryptage/Décryptage à la volée des fichiers images (avec un mot clé que vous devrez changer pour plus de sécurité)

Production :
----------------------
* La mise en production est actuellement faite à l'aide du logiciel InnoSetup.
* Mécanisme de production d'une version en développement :
	* Instrumente le code
	* Crypte les images
	* Compile
	* Crée un fichier de version (utilisé par l'updater)

Instrumentation de code :
----------------------
* Permet d'obtenir des rapports détaillés d'erreur
* Le programme peut détecter l'erreur, avertir l'utilisateur via une fenêtre de type "Oups" et proposer l'envoi du rapport sur votre serveur

Updater intelligent :
----------------------
* Permet de gérer l'espace du jeu en fonction du fichier de version présent sur le serveur
* Demande les droits administrateurs lorsque nécessaire afin d'enregistrer des DLL supplémentaires

Comptes :
----------------------
* Sauvegarde périodique dans une base MySQL (via un pool de connexions)
* Comptes utilisateurs associés à des comptes PHPBB
	
Corrections :
----------------------
* Reprise DirectX suite à un verrouillage session
* Correction de nombreux bugs
	
Note :
Du code permettant de faire du transfert UDP avec utilisation de l'UPNP a été retiré car abandonné dans la version courante.
N'hésitez pas à demander le prototype si besoin car il était parfaitement fonctionnel.
