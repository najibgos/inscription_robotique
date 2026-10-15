# 🤖 استمارة التسجيل في ورشات الروبوتيك — Formulaire Streamlit

Application **Streamlit** de gestion des inscriptions aux ateliers de formation en robotique,
destinée aux élèves du **tronc commun**.

**Entête officielle :** المديرية الإقليمية نواصر — الثانوية التأهيلية علال بن عبد الله
(Direction Provinciale Nouacer — Lycée Qualifiant Alal Ben Abdellah)

## ✨ Fonctionnalités

- Interface entièrement en **arabe (RTL)** avec police arabe (Cairo) et champs alignés à droite
- Entête officielle de l'établissement
- Champs du formulaire : **الاسم العائلي** (nom), **الاسم الشخصي** (prénom), **الفوج** (choix parmi 4 groupes)
- **Limite stricte : 20 élèves maximum par groupe**, contrôlée côté application au moment de la validation
- Insistance sur la limite des places :
  - bandeau d'avertissement (⚠️ عدد المقاعد محدود) ;
  - bandeau de statistiques en temps réel (capacité totale 80, places restantes, inscriptions, groupes complets) ;
  - cartes d'état par groupe avec barre de progression (orange s'il reste ≤ 5 places, rouge « مكتمل » si complet) ;
  - chaque option du menu déroulant affiche le nombre de places restantes ;
  - message de confirmation indiquant les places restantes après inscription ;
  - bandeau « باب التسجيل مغلق » quand les 4 groupes sont complets.
- Détection des doublons (même nom + prénom déjà inscrit)
- Horodatage des inscriptions à l'heure du Maroc (Africa/Casablanca)
- Sauvegarde automatique dans `inscriptions.json`

## 🚀 Installation & lancement

```bash
pip install streamlit
streamlit run app.py
```

L'application s'ouvre automatiquement sur http://localhost:8501

> Astuce réseau : pour rendre le formulaire accessible depuis les téléphones des élèves
> sur le réseau de l'établissement : `streamlit run app.py --server.address 0.0.0.0`
> puis partager l'adresse `http://IP-du-poste:8501`.

## 📁 Fichiers

| Fichier | Rôle |
|---|---|
| `app.py` | Application Streamlit (formulaire) |
| `inscriptions.json` | Données des inscriptions — créé automatiquement au premier enregistrement |

## 🔄 Réinitialiser les inscriptions

Supprimez le fichier `inscriptions.json` (ou remplacez son contenu par `[]`) puis relancez l'application.

## 📊 Consulter / exploiter les inscriptions

Le fichier `inscriptions.json` contient, pour chaque élève : `id`, `nom`, `prenom`, `groupe`, `date_inscription`.
Il peut être ouvert avec n'importe quel éditeur de texte ou importé dans Excel pour l'impression des listes.
