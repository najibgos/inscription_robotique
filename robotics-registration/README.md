# 🤖 استمارة التسجيل في ورشات الروبوتيك — Formulaire Streamlit

Application **Streamlit** de gestion des inscriptions aux ateliers de formation en robotique,
destinée aux élèves du **tronc commun**.

**Entête officielle :** المديرية الإقليمية نواصر — الثانوية التأهيلية علال بن عبد الله
(Direction Provinciale Nouacer — Lycée Qualifiant Alal Ben Abdellah)

## ✨ Fonctionnalités

- Interface entièrement en **arabe (RTL)** avec police arabe (Cairo) et champs alignés à droite
- Entête officielle de l'établissement
- Champs du formulaire : **الاسم العائلي** (nom), **الاسم الشخصي** (prénom), **القسم** (classe), **الفوج** (choix parmi 4 groupes)
- **Champ classe (القسم)** avec les codes officiels : `TCLSH-1`, `TCSF-1` → `TCSF-9`
  (menu déroulant obligatoire, choix vide refusé à la validation)
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
- **Export Excel** :
  - sauvegarde locale automatique de `inscriptions.xlsx` après chaque inscription (quand l'app tourne sur un ordinateur) ;
  - **espace d'administration** (section « 🔒 فضاء الإدارة » en bas de page, protégé par code) : tableau des inscriptions + bouton de téléchargement `inscriptions_robotique_YYYY-MM-DD.xlsx` ;
  - le classeur contient 2 feuilles en droite-à-gauche : **التسجيلات** (liste complète stylée, en-têtes figés, prête à imprimer) et **إحصائيات** (répartition par فوج avec places restantes et taux de remplissage + répartition par قسم).

## 🚀 Installation & lancement

```bash
pip install -r requirements.txt
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
| `excel_export.py` | Génération du classeur Excel (liste + statistiques) |
| `requirements.txt` | Dépendances (`streamlit`, `openpyxl`) — utilisée par Streamlit Cloud |
| `inscriptions.json` | Données des inscriptions — créé automatiquement au premier enregistrement |
| `inscriptions.xlsx` | Export Excel local — mis à jour automatiquement après chaque inscription |

## 🔐 Espace d'administration (export Excel)

En bas de la page, la section **« 🔒 فضاء الإدارة »** demande un code :

- Code par défaut : **`robot2026`**
- Pour le changer : modifiez la ligne `ADMIN_CODE = "robot2026"` au début de `app.py` (conseillé avant la mise en ligne).
- Une fois le code saisi : tableau des inscrits + bouton **« ⬇️ تحميل قائمة التسجيلات (ملف Excel) »**.

## 🔄 Réinitialiser les inscriptions

Supprimez le fichier `inscriptions.json` (ou remplacez son contenu par `[]`) puis relancez l'application.

## 📊 Consulter / exploiter les inscriptions

Trois méthodes :

1. **Espace d'administration** dans l'app (code `robot2026`) → téléchargement direct du fichier Excel ;
2. **`inscriptions.xlsx`** : mis à jour automatiquement après chaque inscription quand l'app tourne localement ;
3. **`inscriptions.json`** : contient pour chaque élève `id`, `nom`, `prenom`, `classe`, `groupe`, `date_inscription` — ouvrable dans un éditeur de texte ou importable dans Excel.
