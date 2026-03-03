# Publipostage Sécurisé (Word & PDF) pour Grist
Widget Grist permettant une modification rapide de documents word en remplaçant des {{variables}} par le contenu de tables de données, ainsi que leur téléchargement Word/PDF.

# 🏛️ Documentation du Widget

Bienvenue sur le manuel d'utilisation du nouvel outil de génération de documents de l'UD ! 
## 📜Préface
Ce module a été minutieusement façonné pour nous permettre de générer des courriers et des actes administratifs de manière automatisée, simultanée et massive, tout en respectant scrupuleusement la charte graphique de l'État et les impératifs de sécurité de notre administration. 
Cette démarche s'insère dans la mise en oeuvre de la dématérialisation à l'UD, et vise à profiter de la flexibilité et de la puissance de Grist pour s'offrir un petit bonus qui, je l'espère, saura se révéler bien utile pour de futures actions à grande échelle, mais également dans le cadre de la gestion de crise.
> [!TIP]
> Je suis disponible pour le faire évoluer au fil de futurs besoins, mais il est conçu pour que vous puissiez en conserver la jouissance même dans l'éventualité de mon départ, *d'où cette documentation la plus exhaustive possible*.

---

## 🛡️ Un mot sur la sécurité et la confidentialité de vos données

Il est légitime, lorsqu'on manipule des données professionnelles ou personnelles, de s'enquérir de leur sécurité 😉. **Je tiens ici à vous rassurer de la manière la plus formelle : ce module est d'une étanchéité absolue.** <sub>(*je sais qu'on disait pareil du Titanic, mais là c'est vrai !*)</sub>

Contrairement aux outils de publipostage externes qui expédient vos informations on ne sait trop où, cet outil fonctionne selon un principe de **traitement strictement local**. 
* L'intégralité des opérations (lecture du document Word, remplacement des variables, génération des PDF) s'effectue **exclusivement au sein de la mémoire vive de votre propre navigateur web**. 
* **Aucune donnée**, qu'il s'agisse des noms de vos correspondants, des adresses ou du contenu de vos modèles, ne quitte votre ordinateur ni ne transite par Internet. 
* Le code du widget a été réalisé par mes soins, et s'appuie sur des bibliothèques OpenSource fiables et connues[^1], intégrées localement dans l'outil (aucun appel à une technologie extérieure n'est réalisé).
* Le widget est hébergé sur GitHub[^2], de façon sécurisée, et ne peut être modifié que par moi-même. Il ne collecte ni ne stocke aucune donnée : rien ne quitte votre ordinateur !

> [!IMPORTANT]
> Vous pouvez donc opérer ce module en toute quiétude, même pour le traitement d'informations sensibles.

---

## ✒️ Comment forger votre modèle Word ?

Afin que le module puisse accomplir son office, il convient évidemment de lui indiquer les emplacements exacts où insérer les informations issues de la base de données.

Dans votre document Word, encadrez simplement le nom de la variable souhaitée par de doubles accolades.<br>

*Exemple :* `Nous vous prions d'agréer, Monsieur {{NOM}}, l'expression de nos salutations distinguées.`
> [!WARNING]
> non testé pour l'instant avec les formats Libre Office / OpenOffice !

**Règles de bon usage :**
1. L'orthographe doit être d'une rigueur absolue : `{{Prenom}}` n'est pas équivalent à `{{prenom}}`. Le module est sensible aux majuscules et aux minuscules, ainsi qu'aux accents.
2. <ins>Assurez-vous que les variables invoquées existent bel et bien en tant qu'intitulés de colonnes</ins> dans nos tables Grist (qu'il s'agisse de la table des Expéditeurs ou de celle des Destinataires).
3. Sauvegardez votre document au format Word "normal" (**.docx**).
> [!TIP]
> La version actuellement déployée sur Grist n'est qu'une base assez simpliste, mais l'outil peut être intégré tel quel dans des tables plus complexes sans problème *(certains ajustements simples pourront être nécessaires mais rien de bien sorcier)*.
---

## 🧭 Mode d'emploi pas-à-pas

L'utilisation du module au sein de Grist a été pensée pour être la plus intuitive possible :

### Étape 1 : Le chargement
Cliquez sur le bouton **"Choisir un fichier"** et sélectionnez votre modèle `.docx` fraîchement préparé. Pressez ensuite le bouton **"Lancer l'analyse du document"**.
> [!NOTE]
> Jusqu'à présent, pour les modèles testés, l'analyse était quasi instantanée. Si ça traîne un peu, n'hésitez pas à me dire !

### Étape 2 : L'inspection des variables
Le module va ausculter votre document et vous dresser un tableau de bord. Il vous indiquera, pour la ligne actuellement sélectionnée dans Grist, si toutes les variables invoquées ont bien été trouvées. 
* Un indicateur vert (✅) vous assure que la donnée est prête à être intégrée au document.
* Un indicateur rouge (❌) vous signale qu'une donnée est manquante dans la base ou n'a pas pu être trouvée. *(pas de bol...)*

### Étape 3 : La génération
Il ne vous reste plus qu'à choisir la portée de votre action :
* **Pour le destinataire sélectionné :** Idéal pour un traitement au cas par cas.
* **Pour TOUS les destinataires :** Le module générera l'ensemble des documents d'un seul trait et vous les livrera pronto sous la forme d'un dossier compressé (Archive ZIP).
> [!TIP]
> Notez que si la génération de multiples documents est très rapide au format Word, elle peut devenir plus longue pour le format pdf.

Enfin, optez pour le format de votre convenance :
* **👁 Prévisualiser :** Ouvre une fenêtre pour contrôler l'aspect visuel du document final avant tout téléchargement. *en cas de souci, cf le dernier chapitre de cette documentation 😉* 
* **📝 Word (.docx) :** Pour générer le document et y apporter d'ultimes retouches manuelles si nécessaire.
* **📄 PDF (.pdf) :** Pour un document figé, prêt à être signé.

> [!NOTE]
> Bien sûr, je suis ouvert aux suggestions tant sur l'aspect fonctionnel que sur l'aspect visuel.

---

## ❓ Résolution des éventuels *(mais tout petits)* désagréments

> [!IMPORTANT]
> Tout le calcul étant fait localement sur votre machine, si vous fermez l'onglet tout s'arrêtera !

### :octocat: Le téléchargement du PDF semble figé (Archive ZIP) :
* La conversion de plusieurs dizaines de documents au format PDF exige un effort substantiel de la part de l'ordinateur *(qui, soyons honnêtes,a déjà facilement du mal qund ce n'est pas le cas)*. Un écran de chargement vous tiendra informé de la progression. Veillez simplement à ne pas fermer la page durant l'opération.
  
### :octocat: La mise en page du PDF diffère légèrement du document original :
Je vous passe le jargon, mais pour générer les pdf efficacement, j'utilise une petite astuce maison. En gros, je ne convertis pas directement le fichier Word en PDF : le navigateur dessine l'image du document *(bien caché pour ne pas faire moche)* puis transforme cette image en PDF. Dit comme ça, ça peut sembler tiré par les cheveux mais en fait c'est beaaaaaaaucoup plus facile et efficace que de manipuler le fichier Word. Résultat, c'est beaucoup plus rapide, et en principe tout aussi précis, mais il peut arriver qu'une mise en page trop complexe cause des soucis.
> [!TIP]
> **Si ça arrive, pas de panique !** Pour la génération au format Word j'utilise un outil différent, donc vous pouvez télécharger en Word, puis convertir en PDF comme vous le faites d'habitude : la mise en forme sera parfaite.
  
### :octocat: La mise en page de la prévisualisation diffère du document original :
C'est exactement la même chose que le souci ci-dessus[^3]. Si la prévisualisation ne ressemble pas au document original, ne perdez pas de temps à le télécharger au format PDF : téléchargez directement au format Word, en principe il n'y aura aucun problème !

### :octocat: Une erreur rouge survient inopinément :
Assurez-vous que vous n'avez pas égaré une accolade dans votre document Word (par exemple `{NOM}}` au lieu de `{{NOM}}`), et/ou passez me voir !

[^1]: Outils utilisés : [FileSaver.min.js](https://www.npmjs.com/package/file-saver) ; [docx-preview.min.js](https://www.npmjs.com/package/docx-preview) ; [docxtemplater.js](https://docxtemplater.com/docs/goals/) ; [html2pdf.bundle.min.js](https://github.com/eKoopmans/html2pdf.js/tree/main) ; [jszip.min.js](https://stuk.github.io/jszip/) ; [pizzip.min.js](https://www.jsdelivr.com/package/npm/pizzip).
[^2]: Pour les curieux : [➡lien vers le GitHub](https://github.com/clement669/Publipostage_Word_Grist).
[^3]: et pour cause : j'utilise le même outil pour dessiner la prévisualisation que pour le "dessin" qui sert de base au pdf !
---
<sub>*Outil développé et maintenu par Clément GAGNOT exclusivement destiné à un usage en interne au sein de l'UD28 DREAL.*</sub>
