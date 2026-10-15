# 🤖 استمارة التسجيل في ورشات الروبوتيك — Formulaire Streamlit

Application **Streamlit** de gestion des inscriptions aux ateliers de formation en robotique,
destinée aux élèves du **tronc commun**.

**Entête officielle :** المديرية الإقليمية نواصر — الثانوية التأهيلية علال بن عبد الله
(Direction Provinciale Nouacer — Lycée Qualifiant Alal Ben Abdellah)

## ✨ Fonctionnalités

- Interface entièrement en **arabe (RTL)** avec police arabe (Cairo) et champs alignés à droite
- Entête officielle de l'établissement
- Champs du formulaire : **الاسم العائلي** (nom), **الاسم الشخصي** (prénom), **القسم** (classe),
  **الفوج** (choix parmi 4 groupes) et **les 2 séances hebdomadaires** — **الحصة الأولى** et
  **الحصة الثانية** (2 menus distincts jour + créneau, avec le nombre de places restantes
  affiché sur chaque option)
- **Champ classe (القسم)** avec les codes officiels : `TCLSH-1`, `TCSF-1` → `TCSF-10`
  (menu déroulant obligatoire, choix vide refusé à la validation)
- **Choix des 2 séances (2 par semaine)** : ateliers à partir du **lundi 26 octobre 2026**,
  du lundi au jeudi entre 14h00 et 18h00, **séances d'une heure**. Chaque élève suit **2 séances
  par semaine : 2 jours différents, dans le même فوج** — et chaque séance peut avoir
  **son propre horaire** (exemple : الاثنين 14:00–15:00 + الأربعاء 16:00–17:00).
  Chaque menu propose les **16 créneaux** (4 jours × 4 horaires) encore disponibles ; le jour
  de la 2e séance doit être **différent** de celui de la 1re (contrôlé à la validation :
  « يجب أن تكون الحصة الثانية في يوم مختلف »). Les créneaux complets (20/20) sont retirés
  automatiquement des 2 menus. Pour ajuster les créneaux, modifiez les constantes
  `DAYS` / `TIME_SLOTS` en début de `app.py` (et les copies identiques dans `excel_export.py`).
- **Limites strictes contrôlées à la validation** : **20 élèves maximum par فوج**
  (constante `MAX_PER_GROUP`) et **20 élèves maximum par حصة** — jour + horaire
  (constante `MAX_PER_SLOT`) ; chaque inscription occupe 1 place dans **chacune** de ses 2 séances.
  En cas de course entre deux inscriptions simultanées, la limite est revérifiée côté serveur
  au moment de la soumission (pour les 2 séances).
- Insistance sur la limite des places :
  - bandeau d'avertissement (⚠️ عدد المقاعد محدود) ;
  - bandeau de statistiques en temps réel (capacité totale 80, places restantes, inscriptions, groupes complets, حصص complètes) ;
  - cartes d'état par groupe avec barre de progression (orange s'il reste ≤ 5 places, rouge « مكتمل » si complet) ;
  - chaque option du menu déroulant affiche le nombre de places restantes ;
  - message de confirmation indiquant les places restantes après inscription ;
  - bandeau « باب التسجيل مغلق » quand les 4 groupes sont complets.
- Détection des doublons (même nom + prénom déjà inscrit)
- Horodatage des inscriptions à l'heure du Maroc (Africa/Casablanca)
- Sauvegarde automatique dans `inscriptions.json`
- **Double sauvegarde locale de secours après chaque inscription** (quand l'app
  tourne sur un ordinateur) pour ne jamais perdre les informations : un fichier
  **Excel** `inscriptions.xlsx` + un fichier **CSV** `inscriptions.csv` contenant
  exactement les mêmes colonnes — si l'un est perdu ou corrompu, l'autre reste
  exploitable. Le CSV est encodé en **UTF-8 avec BOM** (l'arabe s'affiche
  directement dans Excel) avec séparateur « **;** » (s'ouvre en colonnes dans
  Excel en configuration régionale française/Maroc) et fins de ligne CRLF.
- **Export Excel + CSV** :
  - sauvegarde locale automatique de `inscriptions.xlsx` et `inscriptions.csv` après chaque inscription ;
  - **espace d'administration** (section « 🔒 فضاء الإدارة » en bas de page, protégé par code) : tableau des inscriptions + bouton de téléchargement `inscriptions_robotique_YYYY-MM-DD.xlsx` **+ bouton de téléchargement `inscriptions_robotique_YYYY-MM-DD.csv`** ;
  - le classeur contient 2 feuilles en droite-à-gauche : **التسجيلات** (liste complète stylée avec
    les 2 séances : يوم + توقيت de chacune, en-têtes figés, prête à imprimer) et **إحصائيات**
    (répartition par فوج avec places restantes et taux de remplissage + répartition par قسم +
    **التوزيع حسب الحصة** : jour × créneau avec capacité, places restantes et taux de remplissage +
    **التوزيع حسب أيام الحصتين** : les 6 paires de jours, ordre inversé compté ensemble).

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
| `excel_export.py` | Génération du classeur Excel (liste + statistiques) et du CSV de secours |
| `requirements.txt` | Dépendances (`streamlit`, `openpyxl`) — utilisée par Streamlit Cloud |
| `inscriptions.json` | Données des inscriptions — créé automatiquement au premier enregistrement |
| `inscriptions.xlsx` | Export Excel local — mis à jour automatiquement après chaque inscription |
| `inscriptions.csv` | Sauvegarde CSV de sécurité — mêmes colonnes, mise à jour en même temps que l'Excel |

## 🔐 Espace d'administration (export Excel & CSV)

En bas de la page, la section **« 🔒 فضاء الإدارة »** demande un code :

- Code par défaut : **`robot2026`**
- Pour le changer : modifiez la ligne `ADMIN_CODE = "robot2026"` au début de `app.py` (conseillé avant la mise en ligne).
- Une fois le code saisi : tableau des inscrits + bouton **« ⬇️ تحميل قائمة التسجيلات (ملف Excel) »**
  et bouton **« ⬇️ تحميل قائمة التسجيلات (ملف CSV) »**.

## 🔄 Réinitialiser les inscriptions

Supprimez le fichier `inscriptions.json` (ou remplacez son contenu par `[]`) puis relancez l'application.

## 📊 Consulter / exploiter les inscriptions

Trois méthodes :

1. **Espace d'administration** dans l'app (code `robot2026`) → téléchargement direct du fichier Excel **ou du fichier CSV** ;
2. **`inscriptions.xlsx`** ou **`inscriptions.csv`** : mis à jour automatiquement après chaque inscription quand l'app tourne localement (le CSV s'ouvre directement dans Excel, l'arabe s'affiche correctement grâce au BOM UTF-8) ;
3. **`inscriptions.json`** : contient pour chaque élève `id`, `nom`, `prenom`, `classe`, `groupe`,
   `jour1`, `horaire1`, `jour2`, `horaire2`, `date_inscription` — ouvrable dans un éditeur de
   texte ou importable dans Excel. Les anciens enregistrements (`jour` seul, ou `jour1`/`jour2`
   avec un `horaire` unique partagé) restent lus et comptés correctement dans l'application
   et les statistiques.
