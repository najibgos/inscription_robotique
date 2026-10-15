# -*- coding: utf-8 -*-
"""
استمارة التسجيل في ورشات التكوين في الروبوتيك
المديرية الإقليمية نواصر — الثانوية التأهيلية علال بن عبد الله
=================================================================
Application Streamlit : formulaire d'inscription en arabe (RTL)
- Champs : nom, prénom, classe (القسم), groupe/fouj (الفوج) et الموعد
  (menu unique jour + créneau d'une heure, avec places restantes en temps réel)
- Classes disponibles : TCLSH-1 + TCSF-1 → TCSF-9 (codes Massar)
- Horaires : ateliers à partir du lundi 26 octobre 2026, du lundi au jeudi
  entre 14h00 et 18h00 (séances d'une heure : 14:00-15:00, 15:00-16:00,
  16:00-17:00, 17:00-18:00)
- 4 groupes (فوج) — 20 élèves maximum par groupe
- Créneaux (موعد = jour + horaire) — 20 élèves maximum par créneau
- Insistance sur la limite des places (bandeau d'alerte, statistiques, cartes)
- Sauvegarde automatique des inscriptions dans inscriptions.json
- Export Excel : sauvegarde locale inscriptions.xlsx après chaque inscription
  + espace d'administration protégé par code (aperçu + téléchargement .xlsx)
"""

import json
import os
import sys
from datetime import datetime

import streamlit as st

# Import du module d'export Excel (même dossier que app.py)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
try:
    from excel_export import build_excel_bytes, save_excel_file

    EXCEL_OK = True
except Exception:
    EXCEL_OK = False

try:
    from zoneinfo import ZoneInfo
    TZ = ZoneInfo("Africa/Casablanca")  # heure locale du Maroc
except Exception:
    TZ = None

# ============================================================
# 1) CONFIGURATION GÉNÉRALE
# ============================================================
st.set_page_config(
    page_title="استمارة التسجيل في ورشات الروبوتيك",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="collapsed",
)

MAX_PER_GROUP = 20  # limite stricte : 20 élèves au maximum par groupe
MAX_PER_SLOT = 20   # limite stricte : 20 élèves au maximum par موعد (jour + créneau)

# Code d'accès à l'espace d'administration (export Excel) — à personnaliser !
ADMIN_CODE = "robot2026"

GROUPS = ["الفوج الأول", "الفوج الثاني", "الفوج الثالث", "الفوج الرابع"]

# Classes du tronc commun (codes officiels) — TCLSH : الآداب والعلوم الإنسانية، TCSF : علمي
CLASSES = [
    "TCLSH-1",
    "TCSF-1",
    "TCSF-2",
    "TCSF-3",
    "TCSF-4",
    "TCSF-5",
    "TCSF-6",
    "TCSF-7",
    "TCSF-8",
    "TCSF-9",
]

# Jours et créneaux des ateliers — à partir du lundi 26 octobre 2026,
# du lundi au jeudi entre 14h00 et 18h00 (séances d'une heure)
DAYS = ["الاثنين", "الثلاثاء", "الأربعاء", "الخميس"]
TIME_SLOTS = [
    "14:00 – 15:00",
    "15:00 – 16:00",
    "16:00 – 17:00",
    "17:00 – 18:00",
]

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "inscriptions.json")


# ============================================================
# 2) FONCTIONS DE DONNÉES (fichier JSON local)
# ============================================================
def load_inscriptions():
    """Charge la liste des inscriptions depuis le fichier JSON local."""
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                rows = json.load(f)
            return rows if isinstance(rows, list) else []
        except (json.JSONDecodeError, OSError):
            return []
    return []


def save_inscriptions(rows):
    """Sauvegarde la liste des inscriptions dans le fichier JSON local."""
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)


def group_counts(rows):
    """Retourne le nombre d'inscrits par groupe."""
    counts = {g: 0 for g in GROUPS}
    for row in rows:
        if row.get("groupe") in counts:
            counts[row["groupe"]] += 1
    return counts


def slot_counts(rows):
    """Retourne le nombre d'inscrits par موعد (jour + créneau)."""
    counts = {}
    for row in rows:
        j = row.get("jour", "")
        h = row.get("horaire", "")
        if j in DAYS and h in TIME_SLOTS:
            counts[(j, h)] = counts.get((j, h), 0) + 1
    return counts


def normalize(value):
    """Nettoie une chaîne : supprime les espaces multiples et les extrémités."""
    return " ".join((value or "").split()).strip()


def next_id(rows):
    """Prochain identifiant d'inscription."""
    return max([r.get("id", 0) for r in rows], default=0) + 1


def now_str():
    """Horodatage à l'heure locale du Maroc."""
    return (
        datetime.now(TZ).strftime("%Y-%m-%d %H:%M:%S")
        if TZ
        else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    )


# État courant (rechargé à chaque interaction)
data = load_inscriptions()
counts = group_counts(data)
total_capacity = MAX_PER_GROUP * len(GROUPS)
remaining_total = total_capacity - sum(counts.values())
full_groups = sum(1 for g in GROUPS if counts[g] >= MAX_PER_GROUP)

# État des créneaux (موعد = jour + créneau) avec limite MAX_PER_SLOT
slot_counts_map = slot_counts(data)
all_slots = [(j, h) for j in DAYS for h in TIME_SLOTS]
available_slots = [s for s in all_slots if slot_counts_map.get(s, 0) < MAX_PER_SLOT]
full_slots = len(all_slots) - len(available_slots)

# ============================================================
# 3) MISE EN PAGE : arabe RTL + police arabe + styles
# ============================================================
st.markdown(
    """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap');

        html, body, [class*="css"] * {
            font-family: 'Cairo', 'Noto Naskh Arabic', 'Segoe UI', Tahoma, sans-serif !important;
        }

        /* Conteneur principal : direction droite-à-gauche */
        .block-container {
            direction: rtl;
            text-align: right;
            padding-top: 2.2rem;
            max-width: 880px;
        }

        /* Titres de sections Markdown (h1-h6) alignés à droite
           — le h1 du titre principal garde son centrage (style inline) */
        [data-testid="stMarkdownContainer"] h1,
        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        [data-testid="stMarkdownContainer"] h4,
        [data-testid="stMarkdownContainer"] h5,
        [data-testid="stMarkdownContainer"] h6 {
            direction: rtl;
            text-align: right;
        }

        /* Champs de saisie en RTL */
        .stTextInput input {
            direction: rtl;
            text-align: right;
        }
        .stSelectbox [data-baseweb="select"] > div {
            direction: rtl;
            text-align: right;
        }
        [data-baseweb="popover"] {
            direction: rtl;
            text-align: right;
        }

        /* Masquer les éléments Streamlit inutiles */
        #MainMenu { visibility: hidden; }
        footer { visibility: hidden; }
        header[data-testid="stHeader"] { display: none; }
        [data-testid="stSidebar"] { display: none; }

        /* Entête officielle de l'établissement */
        .official-header {
            border: 2px solid #16495c;
            border-radius: 14px;
            background: linear-gradient(180deg, #f6fbfd 0%, #e8f3f7 100%);
            padding: 16px 20px;
            text-align: center;
            direction: rtl;
        }

        /* Bandeau statistiques */
        .stat-strip { display: flex; gap: 10px; direction: rtl; margin: 14px 0 4px 0; }
        .stat-box {
            flex: 1;
            background: #ffffff;
            border: 1.5px solid #d5e3e8;
            border-radius: 12px;
            padding: 10px 12px;
            text-align: center;
        }
        .stat-num { font-size: 1.45rem; font-weight: 900; color: #0f4c5c; line-height: 1.25; }
        .stat-label { font-size: 0.82rem; color: #607d8b; }

        /* Cartes des groupes */
        .group-card {
            background: #ffffff;
            border: 1.5px solid #d5e3e8;
            border-radius: 12px;
            padding: 12px 14px;
            direction: rtl;
        }
        .bar-bg { background: #eceff1; border-radius: 8px; height: 10px; margin-top: 8px; overflow: hidden; }
        .bar-fill { height: 100%; border-radius: 8px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 4) ENTÊTE OFFICIELLE
# ============================================================
st.markdown(
    """
    <div class="official-header">
        <div style="font-size:1.0rem; font-weight:700; color:#16495c;">المديرية الإقليمية نواصر</div>
        <div style="font-size:1.28rem; font-weight:900; color:#0d3a49; margin-top:4px;">الثانوية التأهيلية علال بن عبد الله</div>
        <div style="width:130px; height:3px; background:#1f7a8c; border-radius:3px; margin:10px auto 0 auto;"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Titre de la page
st.markdown(
    """
    <h1 style="text-align:center; color:#0d3a49; font-size:1.7rem; font-weight:900; margin:16px 0 2px 0;">
        🤖 استمارة التسجيل في ورشات التكوين في الروبوتيك
    </h1>
    <p style="text-align:center; color:#52707c; font-size:1.02rem; margin:0 0 12px 0;">
        مفتوحة لتلاميذ الجذع المشترك
    </p>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# 5) INSISTER SUR LA LIMITE DES PLACES
# ============================================================
st.warning(
    "**⚠️ عدد المقاعد محدود:** لا يتجاوز عدد المقبولين في كل فوج **20 تلميذا** فقط. "
    "تُعتمد التسجيلات حسب **أسبقية الوصول**، ويُغلق الفوج نهائيا بمجرد اكتماله."
)

# Bandeau de statistiques en temps réel
st.markdown(
    f"""
    <div class="stat-strip">
        <div class="stat-box">
            <div class="stat-num">{total_capacity}</div>
            <div class="stat-label">الطاقة الاستيعابية (4 أفواج × 20)</div>
        </div>
        <div class="stat-box">
            <div class="stat-num" style="color:{'#c62828' if remaining_total == 0 else '#2e7d32'};">{remaining_total}</div>
            <div class="stat-label">المقاعد المتبقية</div>
        </div>
        <div class="stat-box">
            <div class="stat-num">{len(data)}</div>
            <div class="stat-label">عدد التسجيلات</div>
        </div>
        <div class="stat-box">
            <div class="stat-num" style="color:{'#c62828' if full_groups > 0 else '#0f4c5c'};">{full_groups}</div>
            <div class="stat-label">أفواج مكتملة</div>
        </div>
        <div class="stat-box">
            <div class="stat-num" style="color:{'#c62828' if full_slots > 0 else '#0f4c5c'};">{full_slots}</div>
            <div class="stat-label">مواعيد مكتملة (من {len(all_slots)})</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Cartes d'état des 4 groupes (grille 2 x 2)
st.markdown("#### 📊 حالة الأماكن في كل فوج")
for start in range(0, len(GROUPS), 2):
    cols = st.columns(2)
    for j, g in enumerate(GROUPS[start : start + 2]):
        registered = counts[g]
        left = MAX_PER_GROUP - registered
        pct = int(registered / MAX_PER_GROUP * 100)
        if left == 0:
            color, status = "#c62828", "مكتمل ✋ — التسجيل مغلق"
        elif left == 1:
            color, status = "#ef6c00", "بقي مقعد واحد فقط! سارع بالتسجيل"
        elif left <= 5:
            color, status = "#ef6c00", f"المتبقي: {left} فقط — سارع بالتسجيل"
        else:
            color, status = "#2e7d32", f"المقاعد المتبقية: {left}"
        with cols[j]:
            st.markdown(
                f"""
                <div class="group-card">
                    <div style="font-weight:800; color:#0d3a49; font-size:1.02rem;">{g}</div>
                    <div style="margin-top:4px; color:#455a64; font-size:0.92rem;">
                        المسجلون: <b>{registered} / {MAX_PER_GROUP}</b>
                    </div>
                    <div class="bar-bg"><div class="bar-fill" style="width:{pct}%; background:{color};"></div></div>
                    <div style="margin-top:7px; font-size:0.88rem; font-weight:800; color:{color};">{status}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
st.caption("تُحدَّث الأعداد تلقائيا بعد كل عملية تسجيل.")

# ============================================================
# 6) FORMULAIRE D'INSCRIPTION
# ============================================================
st.markdown("#### 📋 معلومات التلميذ(ة)")

available_groups = [g for g in GROUPS if counts[g] < MAX_PER_GROUP]

if not available_groups:
    st.error(
        "🔒 **انتهت جميع المقاعد!** تم اكتمال كل الأفواج (4 أفواج × 20 تلميذا)، "
        "وعليه تم **إغلاق باب التسجيل**. نشكر لكم اهتمامكم الكبير."
    )
elif not available_slots:
    st.error(
        "🔒 **جميع المواعيد مكتملة!** وصل العدد إلى الحد الأقصى في كل الأيام والتوقيتات "
        "المعروضة. المرجو الاتصال بإدارة الثانوية لمعرفة ما إذا كانت ستُفتح مواعيد جديدة."
    )
else:
    with st.form("registration_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("الاسم العائلي (اللقب) *", placeholder="أدخل الاسم العائلي...")
        with col2:
            prenom = st.text_input("الاسم الشخصي *", placeholder="أدخل الاسم الشخصي...")

        col3, col4 = st.columns(2)
        with col3:
            classe = st.selectbox(
                "القسم *",
                CLASSES,
                index=None,
                placeholder="اختر قسمك...",
            )
        with col4:
            options = [
                f"{g} — المقاعد المتبقية: {MAX_PER_GROUP - counts[g]}" for g in available_groups
            ]
            choice = st.selectbox("الفوج المطلوب *", options)
        groupe = available_groups[options.index(choice)]

        st.markdown(
            '<p style="font-size:0.82rem; color:#78909c; direction:ltr; text-align:left;">'
            "TCLSH : جذع مشترك الآداب والعلوم الإنسانية &nbsp;|&nbsp; TCSF : جذع مشترك علمي"
            "</p>",
            unsafe_allow_html=True,
        )

        # --- الموعد (jour + créneau d'une heure) : menu unique avec places restantes ---
        slot_labels = []
        slot_map = []
        for j, h in available_slots:
            rem = MAX_PER_SLOT - slot_counts_map.get((j, h), 0)
            # \u2066/\u2069 : isolats LTR pour afficher la plage horaire dans le bon ordre (bidi)
            slot_labels.append(f"يوم {j} — \u2066{h}\u2069 — المقاعد المتبقية: {rem}")
            slot_map.append((j, h))
        slot_choice = st.selectbox(
            "الموعد *",
            slot_labels,
            index=None,
            placeholder="اختر اليوم والتوقيت المناسبين...",
        )

        st.markdown(
            '<p style="font-size:0.85rem; color:#78909c;">'
            "🕐 تبدأ الورشات ابتداء من <b>الاثنين 26 أكتوبر 2026</b> — من الاثنين إلى الخميس، "
            "بين الساعة <b>14:00</b> و<b>18:00</b>، ومدة كل حصة <b>ساعة واحدة</b>. "
            "المرجو اختيار الموعد المناسب لك — <b>عدد المقاعد محدود في كل موعد</b>."
            "</p>",
            unsafe_allow_html=True,
        )

        st.markdown(
            '<p style="font-size:0.85rem; color:#78909c;">'
            "* جميع الحقول إلزامية. تُحفظ معلوماتك في قائمة التسجيلات حسب أسبقية الوصول."
            "</p>",
            unsafe_allow_html=True,
        )

        submitted = st.form_submit_button(
            "✅ تأكيد التسجيل", type="primary", width="stretch"
        )

    # --------------------------------------------------------
    # Traitement de la soumission
    # --------------------------------------------------------
    if submitted:
        errors = []

        nom_clean = normalize(nom)
        prenom_clean = normalize(prenom)
        classe_clean = normalize(classe or "")
        jour_clean, horaire_clean = "", ""

        if len(nom_clean) < 2:
            errors.append("المرجو إدخال الاسم العائلي (حرفان على الأقل).")
        if len(prenom_clean) < 2:
            errors.append("المرجو إدخال الاسم الشخصي (حرفان على الأقل).")
        if classe_clean not in CLASSES:
            errors.append("المرجو اختيار القسم (TCLSH-1 أو TCSF).")

        # الموعد : jour + créneau sélectionnés via le menu unique
        if slot_choice is None or slot_choice not in slot_labels:
            errors.append("المرجو اختيار الموعد (اليوم والتوقيت المناسبين).")
        else:
            jour_clean, horaire_clean = slot_map[slot_labels.index(slot_choice)]
            if jour_clean not in DAYS or horaire_clean not in TIME_SLOTS:
                errors.append("المرجو اختيار الموعد (اليوم والتوقيت المناسبين).")
                jour_clean, horaire_clean = "", ""

        # Recharger les données pour réduire les risques de concurrence
        rows_now = load_inscriptions()
        counts_now = group_counts(rows_now)
        slots_now = slot_counts(rows_now)

        if groupe not in GROUPS or counts_now.get(groupe, 0) >= MAX_PER_GROUP:
            errors.append("عذرا، أصبح هذا الفوج مكتملا للتو. المرجو تحديث الصفحة واختيار فوج آخر.")

        # Limite par créneau (revérifiée au moment de la soumission)
        if (jour_clean, horaire_clean) in all_slots and slots_now.get(
            (jour_clean, horaire_clean), 0
        ) >= MAX_PER_SLOT:
            errors.append(
                f"عذرا، أصبح هذا الموعد مكتملا للتو (حد أقصى {MAX_PER_SLOT} تلاميذا في كل موعد). "
                "المرجو تحديث الصفحة واختيار موعد آخر."
            )

        # Détection des doublons (même nom + prénom)
        full_name = f"{nom_clean} {prenom_clean}"
        for row in rows_now:
            existing = f"{normalize(row.get('nom', ''))} {normalize(row.get('prenom', ''))}"
            if existing == full_name:
                errors.append(
                    "هذا الاسم مسجل مسبقا في القائمة. في حال وجود خطأ، المرجو الاتصال بإدارة الثانوية."
                )
                break

        if errors:
            for e in errors:
                st.error(e)
        else:
            rows_now.append(
                {
                    "id": next_id(rows_now),
                    "nom": nom_clean,
                    "prenom": prenom_clean,
                    "classe": classe_clean,
                    "groupe": groupe,
                    "jour": jour_clean,
                    "horaire": horaire_clean,
                    "date_inscription": now_str(),
                }
            )
            save_inscriptions(rows_now)

            # Sauvegarde Excel locale automatique (silencieuse si FS en lecture seule)
            if EXCEL_OK:
                try:
                    save_excel_file(rows_now, os.path.join(BASE_DIR, "inscriptions.xlsx"),
                                    max_per_group=MAX_PER_GROUP)
                except Exception:
                    pass

            left_after_group = MAX_PER_GROUP - counts_now[groupe] - 1
            left_after_slot = MAX_PER_SLOT - slots_now.get((jour_clean, horaire_clean), 0) - 1

            st.balloons()
            st.success(
                f"🎉 تم تسجيل **{prenom_clean} {nom_clean}** — القسم **{classe_clean}** — في **{groupe}**!\n\n"
                f"🕒 الموعد المختار: يوم **{jour_clean}**، التوقيت **\u2066{horaire_clean}\u2069** "
                "(انطلاقا من الاثنين 26 أكتوبر 2026)\n\n"
                f"عدد المقاعد المتبقية في هذا الفوج: **{left_after_group}** — "
                f"وفي هذا الموعد: **{left_after_slot}**\n\n"
                "ملاحظة: يعتبر التسجيل مؤقتا إلى حين تأكيده من طرف إدارة الثانوية."
            )

# ============================================================
# 8) ESPACE ADMINISTRATION : consultation + export Excel
# ============================================================
st.markdown("#### 🔒 فضاء الإدارة")
with st.expander("فتح فضاء الإدارة — قائمة التسجيلات وتصدير ملف Excel"):
    admin_code = st.text_input("رمز الإدارة:", type="password", placeholder="أدخل رمز الإدارة...")

    if admin_code:
        if admin_code == ADMIN_CODE:
            if not data:
                st.info("لا توجد تسجيلات بعد. ستظهر القائمة هنا بعد أول عملية تسجيل.")
            else:
                if EXCEL_OK:
                    try:
                        import pandas as pd

                        df = pd.DataFrame(
                            [
                                {
                                    "رقم": r.get("id"),
                                    "الاسم العائلي": r.get("nom", ""),
                                    "الاسم الشخصي": r.get("prenom", ""),
                                    "القسم": r.get("classe", ""),
                                    "الفوج": r.get("groupe", ""),
                                    "اليوم": r.get("jour", ""),
                                    "التوقيت": r.get("horaire", ""),
                                    "تاريخ التسجيل": r.get("date_inscription", ""),
                                }
                                for r in data
                            ]
                        )
                        st.dataframe(df, width="stretch", hide_index=True)
                    except Exception:
                        pass

                    stamp = datetime.now(TZ).strftime("%Y-%m-%d") if TZ else datetime.now().strftime("%Y-%m-%d")
                    st.download_button(
                        "⬇️ تحميل قائمة التسجيلات (ملف Excel)",
                        data=build_excel_bytes(data, max_per_group=MAX_PER_GROUP, max_per_slot=MAX_PER_SLOT),
                        file_name=f"inscriptions_robotique_{stamp}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width="stretch",
                    )
                    st.caption(
                        "يحتوي الملف على ورقتين: قائمة التسجيلات كاملة + إحصائيات حسب الفوج والقسم والموعد. "
                        "كما يُحفظ نسخة محلية تلقائية (inscriptions.xlsx) بعد كل عملية تسجيل عند تشغيل التطبيق على حاسوب الشخصي."
                    )
                else:
                    st.error(
                        "مكتبة openpyxl غير مثبتة. المرجو تنفيذ الأمر: pip install openpyxl ثم إعادة تشغيل التطبيق."
                    )
        else:
            st.error("رمز غير صحيح. المرجو المحاولة من جديد.")

# ============================================================
# 9) PIED DE PAGE
# ============================================================
st.markdown(
    """
    <div style="text-align:center; margin-top:22px; padding-top:12px; border-top:1px solid #e0e7ea;
                color:#90a4ae; font-size:0.85rem; direction:rtl;">
        للاستفسار، المرجو التواصل مع إدارة الثانوية — الثانوية التأهيلية علال بن عبد الله
    </div>
    """,
    unsafe_allow_html=True,
)
