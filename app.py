import os
import json
import datetime
from io import BytesIO

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import networkx as nx
from openpyxl import Workbook

from sqlalchemy import create_engine, Column, Integer, Float, String, ForeignKey, Table, MetaData, inspect
from sqlalchemy.orm import declarative_base, sessionmaker, relationship, Session
from sqlalchemy.exc import IntegrityError

from dotenv import load_dotenv

# -----------------------------------------------------------
# 1. Конфигурация Streamlit (должна быть первой командой Streamlit)
# -----------------------------------------------------------
APP_VERSION = "2.0"

st.set_page_config(
    page_title="РискНавигатор",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -----------------------------------------------------------
# 2. Загрузка конфигурации БД
# -----------------------------------------------------------
load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///risk_management.db")

# Для PostgreSQL и других СУБД — проверяем и создаём БД
if not DATABASE_URL.startswith("sqlite"):
    try:
        from sqlalchemy_utils import create_database, database_exists
        if not database_exists(DATABASE_URL):
            create_database(DATABASE_URL)
    except ImportError:
        pass

engine = create_engine(DATABASE_URL, echo=False)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

# -----------------------------------------------------------
# 3. Константы
# -----------------------------------------------------------
PROJECT_STATUSES = ["Планирование", "В процессе", "Завершен", "Приостановлен"]
PROJECT_STAGES = ["Планирование", "Разработка", "Тестирование", "Внедрение", "Сопровождение"]

RISK_SPHERES = [
    "Финансовые риски", "Коммерческие риски", "Производственные риски",
    "Экологические риски", "Риски безопасности", "Социальные риски", "Политические риски"
]
RISK_MANAGEMENT = ["Стратегические риски", "Тактические риски", "Оперативные риски"]
RISK_LEVELS = ["Незначительные риски", "Умеренные риски", "Значительные риски", "Критические риски"]
RISK_DANGER = ["Критические риски", "Недопустимые риски", "Допустимые риски", "Незначительные риски"]
RISK_DURATION = ["Катастрофические риски", "Краткосрочные риски", "Среднесрочные риски", "Долгосрочные риски"]
RISK_INSURANCE = ["Не подлежащие страхованию", "Частично подлежащие страхованию", "Полностью подлежащие страхованию"]
RISK_IMPACT = ["Катастрофические", "Значительные", "Незначительные"]
RISK_ORIGIN = ["Внешние риски", "Внутренние риски"]

COLUMN_TRANSLATIONS = {
    'Asset': 'Актив',
    'Category': 'Категория',
    'Cost': 'Стоимость (₽)',
    'Criticality': 'Критичность',
    'Description': 'Описание',
    'Name': 'Название',
    'Type': 'Тип',
    'Probability': 'Вероятность',
    'Level': 'Уровень',
    'Value': 'Ценность (₽)',
    'Threat': 'Угроза',
    'Control': 'Контроль',
    'Mitigated Risk': 'Сниженный риск',
    'Base Risk': 'Базовый риск',
    'Effectiveness': 'Эффективность',
    'ScenarioID': 'ID сценария',
    'Vulnerability': 'Уязвимость'
}

# -----------------------------------------------------------
# 3b. Методические константы (по диссертации)
# -----------------------------------------------------------
# 15 факторов результативности в 5 категориях (раздел 3.4 / табл. 2.2.15).
# Веса по умолчанию: для четырёх факторов — экспертные оценки влияния,
# приведённые в диссертации (по 10-балльной шкале); остальные заданы
# нейтрально (8.0) и подлежат калибровке пользователем.
MATURITY_FACTORS = {
    "Организационные": [
        ("Поддержка высшего руководства", 9.3),
        ("Интеграция с бизнес-процессами", 8.0),
        ("Ясное распределение ролей", 8.0),
    ],
    "Технологические": [
        ("Автоматизация рутинных операций", 8.0),
        ("Применение готовых шаблонов и инструментов", 8.0),
        ("Масштабируемость решений", 8.0),
    ],
    "Методологические": [
        ("Простота и понятность методологии", 9.5),
        ("Фокус на критических рисках", 9.2),
        ("Итеративный подход к внедрению", 8.0),
        ("Баланс формальных и неформальных процедур", 8.0),
    ],
    "Человеческие": [
        ("Компетенции ключевых сотрудников", 9.1),
        ("Осведомлённость персонала", 8.0),
        ("Культура безопасности", 8.0),
    ],
    "Внешние": [
        ("Доступность внешней экспертизы", 8.0),
        ("Соответствие нормативным требованиям", 8.0),
    ],
}

# Уровни (полосы) зрелости по шкале 0–4 (адаптированная модель зрелости).
MATURITY_BANDS = [
    (0.0, 1.0, "Начальный", "На предприятии нет системного управления рисками ИБ; меры применяются реактивно."),
    (1.0, 2.0, "Развивающийся", "Появляются отдельные практики и документы, но процессы не формализованы."),
    (2.0, 3.0, "Определённый", "Процессы управления рисками описаны и применяются регулярно."),
    (3.0, 4.0, "Управляемый", "Процессы измеряются и оптимизируются; решения опираются на данные."),
    (4.0, 5.01, "Оптимизируемый", "Непрерывное совершенствование; управление рисками встроено в стратегию."),
]

# Классификационная матрица моделей управления рисками ИБ (приложение Б): 6 классов.
MODEL_CLASSES = [
    {
        "code": "I",
        "name": "Корпоративные полноформатные",
        "examples": "NIST SP 800-37/53, COBIT 2019, ISO/IEC 27005, ГОСТ 57580",
        "resource": "Высокая",
        "scalability": "Низкая",
        "formalization": "Высокая",
        "ru_compliance": "Средняя",
        "fit": "≤5% самых зрелых МСП (финтех, облачные провайдеры)",
        # профиль идеального применения для скоринга
        "size": {"микро": 0, "малое": 1, "среднее": 2},
        "maturity": {"низкий": -2, "средний": 0, "высокий": 3},
        "regulation": {"низкая": -1, "средняя": 0, "высокая": 2},
        "budget": {"ограниченный": -3, "средний": -1, "значительный": 3},
        "expertise": {"нет": -3, "частично": -1, "штатный": 2},
        "criticality": {"низкая": -1, "средняя": 0, "высокая": 2},
        "base": 0,
    },
    {
        "code": "II",
        "name": "Отраслевые регуляторные",
        "examples": "PCI DSS v4.0, 14-МР ЦБ РФ, ФСТЭК (187-ФЗ КИИ), HIPAA",
        "resource": "Средняя",
        "scalability": "Средняя",
        "formalization": "Средняя",
        "ru_compliance": "Высокая",
        "fit": "15–20% МСП в строго регулируемых нишах",
        "size": {"микро": 0, "малое": 1, "среднее": 1},
        "maturity": {"низкий": 0, "средний": 1, "высокий": 1},
        "regulation": {"низкая": -3, "средняя": 0, "высокая": 4},
        "budget": {"ограниченный": -1, "средний": 1, "значительный": 1},
        "expertise": {"нет": -1, "частично": 1, "штатный": 1},
        "criticality": {"низкая": -1, "средняя": 1, "высокая": 2},
        "base": 1,
    },
    {
        "code": "III",
        "name": "Легковесные (SME-friendly)",
        "examples": "NIST CSF 2.0 Small Business Quick-Start, ENISA для МСП, UK Cyber Essentials",
        "resource": "Низкая",
        "scalability": "Высокая",
        "formalization": "Низкая",
        "ru_compliance": "Средняя",
        "fit": "Лучший выбор для 50–60% МСП",
        "size": {"микро": 2, "малое": 1, "среднее": 0},
        "maturity": {"низкий": 3, "средний": 1, "высокий": -1},
        "regulation": {"низкая": 1, "средняя": 0, "высокая": -1},
        "budget": {"ограниченный": 3, "средний": 1, "значительный": -1},
        "expertise": {"нет": 2, "частично": 1, "штатный": 0},
        "criticality": {"низкая": 1, "средняя": 0, "высокая": -1},
        "base": 3,
    },
    {
        "code": "IV",
        "name": "Модульные / компонентные",
        "examples": "OCTAVE-Allegro, FAIR-CAM, ISF IRAM2, Tiered RMF",
        "resource": "Переменная",
        "scalability": "Высокая",
        "formalization": "Переменная",
        "ru_compliance": "Средняя",
        "fit": "30–40% МСП, готовых «собрать» решение под себя",
        "size": {"микро": 0, "малое": 2, "среднее": 1},
        "maturity": {"низкий": -1, "средний": 2, "высокий": 1},
        "regulation": {"низкая": 0, "средняя": 1, "высокая": 0},
        "budget": {"ограниченный": 0, "средний": 2, "значительный": 1},
        "expertise": {"нет": -2, "частично": 2, "штатный": 2},
        "criticality": {"низкая": 0, "средняя": 1, "высокая": 1},
        "base": 1,
    },
    {
        "code": "V",
        "name": "Интегрированные ERM-ориентированные",
        "examples": "COSO ERM + ISO 31000, CSF 2.0 + Value-Driven Cybersecurity, DORA-SME",
        "resource": "Средняя",
        "scalability": "Средняя",
        "formalization": "Низкая",
        "ru_compliance": "Средняя",
        "fit": "25–30% МСП со зрелой стратегией управления рисками",
        "size": {"микро": -1, "малое": 1, "среднее": 2},
        "maturity": {"низкий": -2, "средний": 1, "высокий": 3},
        "regulation": {"низкая": 0, "средняя": 1, "высокая": 1},
        "budget": {"ограниченный": -1, "средний": 1, "значительный": 2},
        "expertise": {"нет": -1, "частично": 1, "штатный": 2},
        "criticality": {"низкая": 0, "средняя": 1, "высокая": 2},
        "base": 0,
    },
    {
        "code": "VI",
        "name": "Платформенные (SaaS / AI)",
        "examples": "XDR-Risk-Scoring SaaS, отечественные «SOC-в-облаке» (MaxPatrol SIEM Lite)",
        "resource": "Низкая",
        "scalability": "Высокая",
        "formalization": "Низкая",
        "ru_compliance": "Средняя",
        "fit": "Быстро растущий сегмент (~10%); дополняет классы III–IV",
        "size": {"микро": 1, "малое": 1, "среднее": 1},
        "maturity": {"низкий": 1, "средний": 1, "высокий": 0},
        "regulation": {"низкая": 0, "средняя": 1, "высокая": 0},
        "budget": {"ограниченный": 0, "средний": 2, "значительный": 2},
        "expertise": {"нет": 3, "частично": 1, "штатный": -1},
        "criticality": {"низкая": 1, "средняя": 1, "высокая": 0},
        "base": 1,
    },
]

# -----------------------------------------------------------
# 3c. Палитра, тема графиков и оформление (UI)
# -----------------------------------------------------------
INK = "#16243A"        # основной тёмный
NAVY = "#0F1E33"       # сайдбар
ACCENT = "#2F6FED"     # фирменный акцент (азур)
ACCENT_SOFT = "#E8F0FE"
MUTED = "#5B6B82"
SURFACE = "#FFFFFF"
BG = "#F4F6FB"

# Семантика уровней риска
RISK_COLORS = {
    "low": "#16A34A",
    "moderate": "#F59E0B",
    "high": "#EA580C",
    "critical": "#DC2626",
}
PLOTLY_FONT = "Manrope, 'Segoe UI', system-ui, sans-serif"


def style_fig(fig, height=None):
    """Единое оформление графиков Plotly под тему приложения."""
    fig.update_layout(
        font=dict(family=PLOTLY_FONT, color=INK, size=13),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=40, r=24, t=48, b=40),
        title=dict(font=dict(size=16, color=INK)),
        legend=dict(bgcolor="rgba(255,255,255,0.6)", bordercolor="#E3E8F0", borderwidth=1),
        colorway=[ACCENT, "#0E9F9F", "#7C5CFC", "#F59E0B", "#EA580C", "#16A34A"],
    )
    fig.update_xaxes(gridcolor="#EAEEF5", zerolinecolor="#EAEEF5", linecolor="#D7DEEA")
    fig.update_yaxes(gridcolor="#EAEEF5", zerolinecolor="#EAEEF5", linecolor="#D7DEEA")
    if height:
        fig.update_layout(height=height)
    return fig


def inject_css():
    """Внедрение фирменного оформления (один раз за сессию рендера)."""
    st.markdown(
        f"""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"], .stApp {{
            font-family: {PLOTLY_FONT};
            color: {INK};
        }}
        .stApp {{
            background:
                radial-gradient(1200px 600px at 80% -10%, #EAF0FB 0%, rgba(234,240,251,0) 60%),
                {BG};
        }}
        .block-container {{ padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1280px; }}

        /* Заголовки */
        h1, h2, h3, h4 {{ color: {INK}; font-weight: 700; letter-spacing: -0.01em; }}
        h2 {{ font-size: 1.45rem; }} h3 {{ font-size: 1.15rem; }}

        /* Шапка-баннер */
        .rn-header {{
            display: flex; align-items: center; gap: 16px;
            background: linear-gradient(110deg, {NAVY} 0%, #16335C 55%, #1E4C8A 100%);
            color: #fff; border-radius: 18px; padding: 18px 24px; margin-bottom: 18px;
            box-shadow: 0 10px 30px rgba(15,30,51,0.20);
        }}
        .rn-header .rn-logo {{
            width: 46px; height: 46px; border-radius: 12px; flex: 0 0 auto;
            background: linear-gradient(135deg, {ACCENT} 0%, #16C2C2 100%);
            display: flex; align-items: center; justify-content: center; font-size: 24px;
            box-shadow: inset 0 0 0 1px rgba(255,255,255,0.25);
        }}
        .rn-header h1 {{ color: #fff; font-size: 1.5rem; margin: 0; line-height: 1.1; }}
        .rn-header .rn-sub {{ color: #BFD3F2; font-size: 0.9rem; margin-top: 2px; }}
        .rn-header .rn-ver {{
            margin-left: auto; font-size: 0.75rem; color: #cfe0fb;
            border: 1px solid rgba(255,255,255,0.25); padding: 4px 10px; border-radius: 999px;
        }}

        /* Карточки */
        .rn-card {{
            background: {SURFACE}; border: 1px solid #E6EBF3; border-radius: 16px;
            padding: 18px 20px; box-shadow: 0 4px 16px rgba(16,36,58,0.05); margin-bottom: 14px;
        }}
        .rn-pill {{
            display: inline-block; padding: 3px 12px; border-radius: 999px;
            font-size: 0.78rem; font-weight: 600; background: {ACCENT_SOFT}; color: {ACCENT};
        }}

        /* Метрики */
        [data-testid="stMetric"] {{
            background: {SURFACE}; border: 1px solid #E6EBF3; border-radius: 14px;
            padding: 14px 16px; box-shadow: 0 4px 14px rgba(16,36,58,0.05);
        }}
        [data-testid="stMetricValue"] {{ color: {INK}; font-weight: 700; }}
        [data-testid="stMetricLabel"] {{ color: {MUTED}; }}

        /* Кнопки */
        .stButton > button {{
            border-radius: 10px; font-weight: 600; border: 1px solid #D7DEEA;
            transition: all .15s ease;
        }}
        .stButton > button:hover {{ border-color: {ACCENT}; color: {ACCENT}; }}
        .stButton > button[kind="primary"] {{
            background: {ACCENT}; border-color: {ACCENT}; box-shadow: 0 4px 12px rgba(47,111,237,0.35);
        }}
        .stDownloadButton > button {{ border-radius: 10px; font-weight: 600; }}

        /* Сайдбар */
        [data-testid="stSidebar"] {{
            background: linear-gradient(180deg, {NAVY} 0%, #122440 100%);
        }}
        [data-testid="stSidebar"] * {{ color: #DCE6F6; }}
        [data-testid="stSidebar"] .rn-brand {{
            color: #fff; font-weight: 800; font-size: 1.15rem; padding: 6px 4px 2px 4px;
        }}
        [data-testid="stSidebar"] .rn-brand small {{ display:block; color:#9FB6DA; font-weight:500; font-size:.72rem; }}
        [data-testid="stSidebar"] .rn-navcap {{
            color: #7E97C2; font-size: 0.72rem; font-weight: 700; letter-spacing: .08em;
            text-transform: uppercase; margin: 14px 6px 4px 6px;
        }}
        /* Кнопки навигации в сайдбаре */
        [data-testid="stSidebar"] .stButton > button {{
            width: 100%; text-align: left; justify-content: flex-start;
            background: transparent; border: none; color: #C9D8F0; padding: 7px 12px;
            border-radius: 9px; font-weight: 500;
        }}
        [data-testid="stSidebar"] .stButton > button:hover {{
            background: rgba(255,255,255,0.07); color: #fff;
        }}
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {{
            background: rgba(47,111,237,0.22); color: #fff; box-shadow: none;
            border-left: 3px solid {ACCENT};
        }}

        /* Экспандеры */
        [data-testid="stExpander"] {{
            border: 1px solid #E6EBF3; border-radius: 12px; background: {SURFACE};
        }}
        [data-testid="stExpander"] summary:hover {{ color: {ACCENT}; }}

        /* Вкладки */
        .stTabs [data-baseweb="tab-list"] {{ gap: 6px; }}
        .stTabs [data-baseweb="tab"] {{
            border-radius: 9px 9px 0 0; padding: 8px 16px; font-weight: 600;
        }}
        .stTabs [aria-selected="true"] {{ color: {ACCENT}; }}

        /* Таблицы */
        [data-testid="stDataFrame"] {{ border-radius: 12px; overflow: hidden; border: 1px solid #E6EBF3; }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    st.markdown(
        f"""
        <div class="rn-header">
            <div class="rn-logo">🛡️</div>
            <div>
                <h1>РискНавигатор</h1>
                <div class="rn-sub">Управление информационными рисками в системе экономической безопасности малых предприятий</div>
            </div>
            <div class="rn-ver">версия {APP_VERSION}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def maturity_band(level):
    """Возвращает (название, описание) полосы зрелости по значению уровня."""
    for lo, hi, name, desc in MATURITY_BANDS:
        if lo <= level < hi:
            return name, desc
    return "—", ""

# -----------------------------------------------------------
# 4. Модели ORM (все в одном месте)
# -----------------------------------------------------------
class Risk(Base):
    __tablename__ = "risks"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    sphere = Column(String)
    management = Column(String)
    level = Column(String)
    danger = Column(String)
    duration = Column(String)
    insurance = Column(String)
    impact = Column(String)
    origin = Column(String)


class Project(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True)
    description = Column(String)
    status = Column(String, default="В процессе")
    risks = relationship("RiskRegister", back_populates="project", cascade="all, delete-orphan")


class Expert(Base):
    __tablename__ = "experts"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String)
    position = Column(String)
    evaluations = relationship("RiskEvaluation", back_populates="expert", cascade="all, delete-orphan")


class RiskRegister(Base):
    __tablename__ = "risk_register"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String)
    name = Column(String)
    stage = Column(String)
    project_id = Column(Integer, ForeignKey("projects.id"))

    project = relationship("Project", back_populates="risks")
    evaluations = relationship("RiskEvaluation", back_populates="risk", cascade="all, delete-orphan")


class RiskEvaluation(Base):
    __tablename__ = "risk_evaluations"
    id = Column(Integer, primary_key=True, index=True)
    risk_id = Column(Integer, ForeignKey("risk_register.id"))
    expert_id = Column(Integer, ForeignKey("experts.id"))
    probability = Column(Float)
    impact = Column(Float)
    evaluation_date = Column(String)

    risk = relationship("RiskRegister", back_populates="evaluations")
    expert = relationship("Expert", back_populates="evaluations")


class Asset(Base):
    __tablename__ = "assets"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True)
    category = Column(String)
    value = Column(Float, default=0.0)
    criticality = Column(Float, default=0.5)
    scenarios = relationship("RiskScenario", back_populates="asset")


class Threat(Base):
    __tablename__ = "threats"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True)
    type = Column(String)
    probability = Column(Float, default=0.5)
    description = Column(String)
    scenarios = relationship("RiskScenario", back_populates="threat")


class Vulnerability(Base):
    __tablename__ = "vulnerabilities"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True)
    level = Column(Float, default=0.5)
    description = Column(String)
    scenarios = relationship("RiskScenario", back_populates="vuln")


class Control(Base):
    __tablename__ = "controls"
    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, unique=True)
    type = Column(String)
    effectiveness = Column(Float, default=0.5)
    cost = Column(Float, default=0.0)
    description = Column(String)
    scenarios = relationship("RiskScenario", back_populates="control")


class RiskScenario(Base):
    __tablename__ = "risk_scenarios"
    id = Column(Integer, primary_key=True, index=True)
    asset_id = Column(Integer, ForeignKey("assets.id"))
    threat_id = Column(Integer, ForeignKey("threats.id"))
    vuln_id = Column(Integer, ForeignKey("vulnerabilities.id"))
    control_id = Column(Integer, ForeignKey("controls.id"), nullable=True)

    asset = relationship("Asset", back_populates="scenarios")
    threat = relationship("Threat", back_populates="scenarios")
    vuln = relationship("Vulnerability", back_populates="scenarios")
    control = relationship("Control", back_populates="scenarios")


class RiskEvent(Base):
    __tablename__ = "risk_events"
    id = Column(Integer, primary_key=True, index=True)
    title = Column(String)
    description = Column(String)
    date = Column(String)
    type = Column(String)
    impact = Column(Float, default=0.0)
    probability = Column(Float, default=0.0)


class MaturityAssessment(Base):
    """Результат оценки зрелости процессов управления рисками ИБ."""
    __tablename__ = "maturity_assessments"
    id = Column(Integer, primary_key=True, index=True)
    org_name = Column(String)
    date = Column(String)
    overall_score = Column(Float)       # средневзвешенная оценка по 15 факторам (1..5)
    ml_predicted = Column(Float)        # прогноз по формуле модели зрелости (0..4)
    factor_scores = Column(String)      # JSON: {фактор: оценка 1..5}
    comment = Column(String)


# Создаём все таблицы
Base.metadata.create_all(engine)

# -----------------------------------------------------------
# 5. Функции расчёта
# -----------------------------------------------------------
def calculate_risk(asset_value, threat_probability, vuln_level, control_effectiveness=0.0):
    """Расчёт базового и остаточного риска."""
    if not (0 <= threat_probability <= 1):
        raise ValueError("Вероятность угрозы должна быть [0..1].")
    if not (0 <= vuln_level <= 1):
        raise ValueError("Уровень уязвимости должен быть [0..1].")
    if not (0 <= control_effectiveness <= 1):
        raise ValueError("Эффективность контроля должна быть [0..1].")

    base_risk = asset_value * threat_probability * vuln_level
    mitigated_risk = base_risk * (1 - control_effectiveness)
    return base_risk, mitigated_risk


def calculate_basic_risk(probability, impact):
    """Базовый рейтинг: Вероятность * Влияние"""
    return probability * impact


def calculate_weighted_risk(probability, impact):
    """Взвешенный рейтинг: 40% вероятность + 60% влияние"""
    return (0.4 * probability) + (0.6 * impact)


def get_risk_category(rating):
    """Категорийный рейтинг: числовой → текстовый"""
    if rating <= 4:
        return "Низкий риск", "green"
    elif rating <= 10:
        return "Умеренный риск", "orange"
    elif rating <= 16:
        return "Высокий риск", "red"
    else:
        return "Критический риск", "darkred"


def calculate_complex_risk(probability, impact, additional_factor=1.0):
    """Комплексный рейтинг с дополнительным фактором"""
    return probability * impact * additional_factor


def calculate_roi(cost, benefit):
    """ROI = ((Benefit - Cost) / Cost) * 100"""
    if cost <= 0:
        return None
    return ((benefit - cost) / cost) * 100


def calculate_tco(initial_cost, operational_cost, years):
    """TCO = Initial Cost + (Operational Cost * Years)"""
    return initial_cost + (operational_cost * years)


def calculate_evi(delta_risk, total_cost):
    """
    EVI — индекс эффективности меры защиты (Economic Value of Information-security control).
    Основной экономический показатель методики (см. гл. 2-3 диссертации).
    Сопоставляет ожидаемое снижение риска с совокупными затратами на меру:
        EVI = ΔR / Совокупные затраты,
    где ΔR = Базовый риск − Остаточный риск.
    Интерпретация: сколько единиц снижения риска приходится на единицу затрат.
    EVI ≥ 1 — мера окупается по риску; чем выше, тем приоритетнее мера.
    """
    if total_cost is None or total_cost <= 0:
        return None
    return delta_risk / total_cost


# -----------------------------------------------------------
# 6. Утилиты
# -----------------------------------------------------------
def translate_columns(df):
    """Перевод заголовков столбцов на русский."""
    return df.rename(columns=COLUMN_TRANSLATIONS)


def export_to_excel(dataframes_dict):
    """Экспорт словаря DataFrame-ов в Excel."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in dataframes_dict.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output


def safe_index(lst, value, default=0):
    """Безопасный поиск индекса в списке."""
    try:
        return lst.index(value)
    except (ValueError, AttributeError):
        return default


def display_risk_category(rating):
    """Безопасное отображение категории риска (без unsafe_allow_html)."""
    category, color = get_risk_category(rating)
    if color == "green":
        st.success(f"{category} (рейтинг: {rating:.1f})")
    elif color == "orange":
        st.warning(f"{category} (рейтинг: {rating:.1f})")
    else:
        st.error(f"{category} (рейтинг: {rating:.1f})")


# -----------------------------------------------------------
# 7. Матрица рисков
# -----------------------------------------------------------
def create_risk_matrix(fig_width=600, fig_height=400):
    """Создание матрицы рисков."""
    matrix = np.zeros((5, 5))

    for i in range(5):
        for j in range(5):
            risk_value = (i + 1) * (j + 1)
            if risk_value <= 4:
                matrix[4 - i, j] = 0
            elif risk_value <= 10:
                matrix[4 - i, j] = 1
            elif risk_value <= 16:
                matrix[4 - i, j] = 2
            else:
                matrix[4 - i, j] = 3

    fig = px.imshow(
        matrix,
        labels=dict(x="Влияние →", y="Вероятность →"),
        x=[1, 2, 3, 4, 5],
        y=[5, 4, 3, 2, 1],
        color_continuous_scale=["green", "yellow", "orange", "red"],
        width=fig_width,
        height=fig_height
    )

    fig.update_layout(
        coloraxis_showscale=False,
        margin=dict(l=30, r=30, t=30, b=30)
    )

    return fig


def mark_risk_position(fig, prob, impact):
    """Отметка позиции риска на матрице."""
    prob_idx = min(round(prob), 5) - 1
    impact_idx = min(round(impact), 5) - 1

    fig.add_trace(
        go.Scatter(
            x=[impact_idx + 1],
            y=[5 - prob_idx],
            mode="markers",
            marker=dict(size=12, color="black", symbol="circle"),
            showlegend=False
        )
    )

    return fig


# -----------------------------------------------------------
# 8. Инициализация тестовых данных для реестра рисков
# -----------------------------------------------------------
def add_register_sample_data():
    """Добавление примеров данных для реестра рисков."""
    session = SessionLocal()
    try:
        if session.query(Project).first() is not None:
            return

        project = Project(
            name="Стратегическое планирование 2025",
            description="Проект по стратегическому планированию на 2025 год",
            status="В процессе"
        )
        session.add(project)
        session.flush()

        experts = [
            Expert(name="Иванов А.С.", position="Руководитель HR"),
            Expert(name="Петрова Е.В.", position="Финансовый аналитик"),
            Expert(name="Сидоров К.М.", position="Технический директор")
        ]
        session.add_all(experts)
        session.flush()

        risk = RiskRegister(
            code="R001",
            name="Недостаток квалифицированных кадров",
            stage="Планирование",
            project_id=project.id
        )
        session.add(risk)
        session.flush()

        current_date = datetime.datetime.now().strftime("%d.%m.%Y")

        evaluations = [
            RiskEvaluation(risk_id=risk.id, expert_id=experts[0].id,
                           probability=5.0, impact=4.0, evaluation_date=current_date),
            RiskEvaluation(risk_id=risk.id, expert_id=experts[1].id,
                           probability=4.0, impact=5.0, evaluation_date=current_date),
            RiskEvaluation(risk_id=risk.id, expert_id=experts[2].id,
                           probability=3.0, impact=4.0, evaluation_date=current_date)
        ]
        session.add_all(evaluations)
        session.commit()
    except Exception as e:
        session.rollback()
        st.error(f"Ошибка при добавлении тестовых данных: {str(e)}")
    finally:
        session.close()


# -----------------------------------------------------------
# 9. Страницы приложения
# -----------------------------------------------------------

# ==================== ГЛАВНАЯ ====================
def load_demo_data():
    """
    Загрузка ДЕМОНСТРАЦИОННЫХ данных (активы, угрозы, уязвимости, меры, сценарии).
    Данные являются условными и предназначены только для демонстрации работы
    программы. Они не отражают результаты реального обследования предприятий.
    """
    session = SessionLocal()
    try:
        if session.query(Asset).first() is not None:
            return False, "Данные уже есть — демонстрационный набор не добавлялся."

        assets = [
            Asset(name="[демо] База клиентов (CRM)", category="Данные", value=1500000, criticality=0.9),
            Asset(name="[демо] Бухгалтерская система 1С", category="ПО", value=900000, criticality=0.85),
            Asset(name="[демо] Корпоративная почта", category="Сервис", value=400000, criticality=0.6),
            Asset(name="[демо] Сайт и интернет-магазин", category="Сервис", value=700000, criticality=0.7),
            Asset(name="[демо] Файловый сервер", category="Инфраструктура", value=500000, criticality=0.65),
        ]
        threats = [
            Threat(name="[демо] Фишинг / соц. инженерия", type="Внешняя", probability=0.7,
                   description="Целевые письма сотрудникам"),
            Threat(name="[демо] Вымогательское ПО (ransomware)", type="Внешняя", probability=0.5,
                   description="Шифрование данных с требованием выкупа"),
            Threat(name="[демо] Утечка данных", type="Смешанная", probability=0.45,
                   description="Несанкционированный доступ к ПДн клиентов"),
            Threat(name="[демо] Сбой/отказ оборудования", type="Внутренняя", probability=0.4,
                   description="Выход из строя сервера без резервирования"),
            Threat(name="[демо] Ошибки / инсайдер", type="Внутренняя", probability=0.35,
                   description="Случайные или умышленные действия персонала"),
        ]
        vulns = [
            Vulnerability(name="[демо] Нет MFA", level=0.8, description="Доступ только по паролю"),
            Vulnerability(name="[демо] Нерегулярные резервные копии", level=0.7, description="Бэкапы не проверяются"),
            Vulnerability(name="[демо] Низкая осведомлённость персонала", level=0.75,
                          description="Сотрудники не обучены распознавать атаки"),
            Vulnerability(name="[демо] Устаревшее ПО", level=0.6, description="Не устанавливаются обновления"),
            Vulnerability(name="[демо] Слабое разграничение прав", level=0.55, description="Избыточные права доступа"),
        ]
        controls = [
            Control(name="[демо] Внедрение MFA", type="Техническая", effectiveness=0.7, cost=60000,
                    description="Двухфакторная аутентификация"),
            Control(name="[демо] Резервное копирование 3-2-1", type="Техническая", effectiveness=0.75, cost=90000,
                    description="Регулярные проверяемые бэкапы off-site"),
            Control(name="[демо] Обучение «anti-phish»", type="Организационная", effectiveness=0.5, cost=40000,
                    description="Тренинги и тестовые рассылки"),
            Control(name="[демо] EDR-защита рабочих станций", type="Техническая", effectiveness=0.65, cost=120000,
                    description="Антивирус нового поколения"),
            Control(name="[демо] Политика управления доступом", type="Организационная", effectiveness=0.45, cost=25000,
                    description="Принцип наименьших привилегий"),
        ]
        session.add_all(assets + threats + vulns + controls)
        session.flush()

        pairs = [(0, 0, 2, 2), (0, 2, 0, 0), (1, 1, 1, 1), (3, 0, 2, 2), (4, 3, 1, None), (1, 4, 4, 4)]
        for ai, ti, vi, ci in pairs:
            session.add(RiskScenario(
                asset_id=assets[ai].id, threat_id=threats[ti].id,
                vuln_id=vulns[vi].id,
                control_id=controls[ci].id if ci is not None else None,
            ))
        session.commit()
        return True, "Демонстрационные данные загружены."
    except Exception as e:
        session.rollback()
        return False, f"Ошибка загрузки демо-данных: {e}"
    finally:
        session.close()


def show_home_page():
    session = SessionLocal()
    try:
        n_assets = session.query(Asset).count()
        n_threats = session.query(Threat).count()
        n_vulns = session.query(Vulnerability).count()
        n_controls = session.query(Control).count()
        scenarios = session.query(RiskScenario).all()
        total_base = total_mitig = 0.0
        for s in scenarios:
            if s.asset and s.threat and s.vuln:
                b, m = calculate_risk(s.asset.value, s.threat.probability, s.vuln.level,
                                      s.control.effectiveness if s.control else 0)
                total_base += b
                total_mitig += m
    finally:
        session.close()

    st.markdown(
        '<div class="rn-card">Программный комплекс реализует процессно-ориентированную методику '
        'управления информационными рисками, адаптированную для малых предприятий: учёт активов, угроз, '
        'уязвимостей и мер контроля, формирование сценариев риска, экономическую приоритизацию мер через '
        '<b>EVI</b>, оценку зрелости процессов и выбор класса модели управления рисками.</div>',
        unsafe_allow_html=True,
    )

    st.markdown("#### Текущее состояние")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Активы", n_assets)
    c2.metric("Угрозы", n_threats)
    c3.metric("Уязвимости", n_vulns)
    c4.metric("Меры контроля", n_controls)

    c5, c6, c7 = st.columns(3)
    c5.metric("Сценарии риска", len(scenarios))
    c6.metric("Суммарный базовый риск", f"{total_base:,.0f} ₽".replace(",", " "))
    reduction = (1 - total_mitig / total_base) * 100 if total_base else 0
    c7.metric("Остаточный риск", f"{total_mitig:,.0f} ₽".replace(",", " "),
              delta=f"−{reduction:.0f}%" if total_base else None, delta_color="inverse")

    st.markdown("")
    left, right = st.columns([1.3, 1])
    with left:
        st.markdown("#### С чего начать")
        st.markdown(
            "1. Внесите **активы** и оцените их ценность и критичность.\n"
            "2. Опишите **угрозы** и **уязвимости**.\n"
            "3. Свяжите их в **сценарии риска** (актив → угроза → уязвимость → контроль).\n"
            "4. На странице **Экономика (EVI, ROI, TCO)** приоритизируйте меры защиты.\n"
            "5. Оцените **зрелость** процессов и при необходимости подберите **класс модели** в мастере."
        )
    with right:
        st.markdown("#### Возможности методики")
        st.markdown(
            "- Мультиуровневая оценка риска (4 метода)\n"
            "- Экспертная оценка и матрица рисков\n"
            "- EVI — экономическая приоритизация мер\n"
            "- Модель зрелости: 15 факторов, 5 категорий\n"
            "- Мастер выбора модели (6 классов)\n"
            "- Тепловые карты, 3D-граф, дашборды, экспорт в Excel"
        )

    with st.expander("Демонстрационные данные"):
        st.caption(
            "Загрузит условный пример (активы, угрозы, уязвимости, меры и сценарии) для ознакомления "
            "с интерфейсом. Данные помечены префиксом «[демо]», являются вымышленными и не отражают "
            "результатов реального обследования предприятий."
        )
        if st.button("Загрузить демонстрационные данные", key="load_demo_btn"):
            ok, msg = load_demo_data()
            (st.success if ok else st.info)(msg)
            if ok:
                st.rerun()


# ==================== РИСКИ ====================
def show_risks_page():
    st.title("Риски")
    session = SessionLocal()

    try:
        risks = session.query(Risk).all()

        df_risks = pd.DataFrame([{
            "ID": r.id,
            "Наименование": r.name,
            "Сфера": r.sphere,
            "Менеджмент": r.management,
            "Уровень": r.level,
            "Опасность": r.danger,
            "Длительность": r.duration,
            "Страхование": r.insurance,
            "Результат": r.impact,
            "Возникновение": r.origin
        } for r in risks])
        st.dataframe(df_risks, width='stretch')

        with st.expander("Добавить риск"):
            name = st.text_input("Наименование риска", key="risk_add_name")
            sphere = st.selectbox("Сфера", RISK_SPHERES, key="risk_add_sphere")
            management = st.selectbox("Менеджмент", RISK_MANAGEMENT, key="risk_add_management")
            level = st.selectbox("Уровень", RISK_LEVELS, key="risk_add_level")
            danger = st.selectbox("Опасность", RISK_DANGER, key="risk_add_danger")
            duration = st.selectbox("Длительность", RISK_DURATION, key="risk_add_duration")
            insurance = st.selectbox("Страхование", RISK_INSURANCE, key="risk_add_insurance")
            impact = st.selectbox("Результат", RISK_IMPACT, key="risk_add_impact")
            origin = st.selectbox("Возникновение", RISK_ORIGIN, key="risk_add_origin")

            if st.button("Сохранить риск", key="risk_add_save"):
                if name:
                    new_risk = Risk(
                        name=name, sphere=sphere, management=management,
                        level=level, danger=danger, duration=duration,
                        insurance=insurance, impact=impact, origin=origin
                    )
                    try:
                        session.add(new_risk)
                        session.commit()
                        st.success("Риск добавлен!")
                        st.rerun()
                    except Exception as e:
                        session.rollback()
                        st.error(f"Ошибка при добавлении риска: {str(e)}")
                else:
                    st.error("Название не может быть пустым")

        with st.expander("Редактировать риск"):
            ids = [r.id for r in risks]
            if ids:
                selected_id = st.selectbox("Выберите ID риска", ids, key="risk_edit_id")
                selected_risk = session.get(Risk, selected_id)

                if selected_risk is None:
                    st.error("Ошибка: риск с таким ID не найден.")
                else:
                    new_name = st.text_input("Наименование риска", value=selected_risk.name, key="risk_edit_name")
                    new_sphere = st.selectbox("Сфера", RISK_SPHERES,
                                              index=safe_index(RISK_SPHERES, selected_risk.sphere),
                                              key="risk_edit_sphere")
                    new_management = st.selectbox("Менеджмент", RISK_MANAGEMENT,
                                                  index=safe_index(RISK_MANAGEMENT, selected_risk.management),
                                                  key="risk_edit_management")
                    new_level = st.selectbox("Уровень", RISK_LEVELS,
                                             index=safe_index(RISK_LEVELS, selected_risk.level),
                                             key="risk_edit_level")
                    new_danger = st.selectbox("Опасность", RISK_DANGER,
                                              index=safe_index(RISK_DANGER, selected_risk.danger),
                                              key="risk_edit_danger")
                    new_duration = st.selectbox("Длительность", RISK_DURATION,
                                                index=safe_index(RISK_DURATION, selected_risk.duration),
                                                key="risk_edit_duration")
                    new_insurance = st.selectbox("Страхование", RISK_INSURANCE,
                                                 index=safe_index(RISK_INSURANCE, selected_risk.insurance),
                                                 key="risk_edit_insurance")
                    new_impact = st.selectbox("Результат", RISK_IMPACT,
                                              index=safe_index(RISK_IMPACT, selected_risk.impact),
                                              key="risk_edit_impact")
                    new_origin = st.selectbox("Возникновение", RISK_ORIGIN,
                                              index=safe_index(RISK_ORIGIN, selected_risk.origin),
                                              key="risk_edit_origin")

                    if st.button("Сохранить изменения", key="risk_edit_save"):
                        try:
                            selected_risk.name = new_name
                            selected_risk.sphere = new_sphere
                            selected_risk.management = new_management
                            selected_risk.level = new_level
                            selected_risk.danger = new_danger
                            selected_risk.duration = new_duration
                            selected_risk.insurance = new_insurance
                            selected_risk.impact = new_impact
                            selected_risk.origin = new_origin
                            session.commit()
                            st.success("Изменения сохранены!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при сохранении: {str(e)}")

        with st.expander("Удалить риск"):
            ids = [r.id for r in risks]
            if ids:
                delete_id = st.selectbox("Выберите ID для удаления", ids, key="risk_delete_id")
                if st.button("Удалить выбранный риск", key="risk_delete_btn"):
                    risk_to_delete = session.get(Risk, delete_id)
                    if risk_to_delete:
                        try:
                            session.delete(risk_to_delete)
                            session.commit()
                            st.success(f"Риск с ID {delete_id} удалён.")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при удалении: {str(e)}")
                    else:
                        st.error("Риск уже удалён или не существует.")
    finally:
        session.close()


# ==================== РЕЕСТР РИСКОВ ====================
def show_risk_register():
    """Отображение реестра рисков."""
    st.title("Экспертная оценка риска")

    add_register_sample_data()

    tab1, tab2, tab3 = st.tabs(["Проекты", "Риски", "Экспертная оценка"])

    with tab1:
        show_projects_tab()
    with tab2:
        show_risks_register_tab()
    with tab3:
        show_expert_evaluation_tab()


def show_projects_tab():
    """Вкладка с управлением проектами."""
    st.subheader("Управление проектами")
    session = SessionLocal()

    try:
        with st.expander("Добавить новый проект", expanded=False):
            name = st.text_input("Название проекта", key="new_project_name")
            description = st.text_area("Описание проекта", key="new_project_desc")
            status = st.selectbox("Статус проекта", PROJECT_STATUSES, key="new_project_status")

            if st.button("Сохранить проект", key="project_save_btn"):
                if name:
                    project = Project(name=name, description=description, status=status)
                    try:
                        session.add(project)
                        session.commit()
                        st.success(f"Проект '{name}' добавлен!")
                        st.rerun()
                    except IntegrityError:
                        session.rollback()
                        st.error("Проект с таким названием уже существует!")
                    except Exception as e:
                        session.rollback()
                        st.error(f"Ошибка при сохранении проекта: {str(e)}")
                else:
                    st.error("Название проекта не может быть пустым")

        projects = session.query(Project).all()

        if projects:
            project_data = []
            for project in projects:
                risk_count = session.query(RiskRegister).filter(RiskRegister.project_id == project.id).count()
                project_data.append({
                    "ID": project.id,
                    "Название": project.name,
                    "Описание": project.description,
                    "Статус": project.status,
                    "Кол-во рисков": risk_count
                })

            df_projects = pd.DataFrame(project_data)
            st.dataframe(df_projects, hide_index=True)

            with st.expander("Редактировать/Удалить проект", expanded=False):
                project_ids = [p.id for p in projects]
                selected_id = st.selectbox("Выберите ID проекта", project_ids, key="edit_project_id")

                if selected_id:
                    project_obj = session.get(Project, selected_id)

                    new_name = st.text_input("Новое название", project_obj.name, key="edit_project_name")
                    new_desc = st.text_area("Новое описание", project_obj.description, key="edit_project_desc")
                    new_status = st.selectbox("Новый статус", PROJECT_STATUSES,
                                              index=safe_index(PROJECT_STATUSES, project_obj.status),
                                              key="edit_project_status")

                    col1, col2 = st.columns(2)
                    if col1.button("Обновить проект", key="project_update_btn"):
                        try:
                            project_obj.name = new_name
                            project_obj.description = new_desc
                            project_obj.status = new_status
                            session.commit()
                            st.success("Проект обновлен!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при обновлении проекта: {str(e)}")

                    if col2.button("Удалить проект", key="project_delete_btn"):
                        try:
                            risk_count = session.query(RiskRegister).filter(
                                RiskRegister.project_id == selected_id).count()
                            if risk_count > 0:
                                st.error(
                                    f"Невозможно удалить проект, так как с ним связано {risk_count} рисков. "
                                    f"Сначала удалите риски.")
                            else:
                                session.delete(project_obj)
                                session.commit()
                                st.success("Проект удален!")
                                st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при удалении проекта: {str(e)}")
        else:
            st.info("Нет проектов. Добавьте новый проект.")
    finally:
        session.close()


def show_risks_register_tab():
    """Вкладка с управлением рисками в реестре."""
    st.subheader("Управление рисками")
    session = SessionLocal()

    try:
        with st.expander("Добавить новый риск", expanded=False):
            projects = session.query(Project).all()

            if not projects:
                st.warning("Сначала создайте проект на вкладке 'Проекты'")
            else:
                risk_count = session.query(RiskRegister).count()
                risk_code = f"R{(risk_count + 1):03d}"

                st.text_input("Код риска", risk_code, disabled=True, key="new_risk_code")
                risk_name = st.text_input("Название риска", key="new_risk_name")

                project_options = {p.name: p.id for p in projects}
                selected_project = st.selectbox("Проект", list(project_options.keys()), key="new_risk_project")

                stage = st.selectbox("Этап", PROJECT_STAGES, key="new_risk_stage")

                if st.button("Сохранить риск", key="register_risk_save_btn"):
                    if risk_name:
                        risk = RiskRegister(
                            code=risk_code,
                            name=risk_name,
                            stage=stage,
                            project_id=project_options[selected_project]
                        )
                        try:
                            session.add(risk)
                            session.commit()
                            st.success(f"Риск '{risk_name}' добавлен!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при сохранении риска: {str(e)}")
                    else:
                        st.error("Название риска не может быть пустым")

        risks = session.query(RiskRegister).all()
        projects = session.query(Project).all()

        if risks:
            risk_data = []
            for risk in risks:
                try:
                    project_name = risk.project.name
                except Exception:
                    project_name = "Неизвестный проект"

                risk_data.append({
                    "ID": risk.id,
                    "Код": risk.code,
                    "Название": risk.name,
                    "Проект": project_name,
                    "Этап": risk.stage,
                    "Кол-во оценок": len(risk.evaluations) if risk.evaluations else 0
                })

            df_risks = pd.DataFrame(risk_data)
            st.dataframe(df_risks, hide_index=True)

            with st.expander("Редактировать/Удалить риск", expanded=False):
                risk_ids = [r.id for r in risks]
                selected_risk_id = st.selectbox("Выберите ID риска", risk_ids, key="edit_risk_id")

                if selected_risk_id:
                    risk_obj = session.get(RiskRegister, selected_risk_id)

                    new_name = st.text_input("Новое название", risk_obj.name, key="edit_risk_name")

                    project_options = {p.name: p.id for p in projects}
                    current_project = next((p.name for p in projects if p.id == risk_obj.project_id), None)
                    project_index = safe_index(list(project_options.keys()), current_project)

                    new_project = st.selectbox("Новый проект", list(project_options.keys()),
                                               index=project_index, key="edit_risk_project")
                    new_stage = st.selectbox("Новый этап", PROJECT_STAGES,
                                             index=safe_index(PROJECT_STAGES, risk_obj.stage),
                                             key="edit_risk_stage")

                    col1, col2 = st.columns(2)
                    if col1.button("Обновить риск", key="register_risk_update_btn"):
                        try:
                            risk_obj.name = new_name
                            risk_obj.project_id = project_options[new_project]
                            risk_obj.stage = new_stage
                            session.commit()
                            st.success("Риск обновлен!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при обновлении риска: {str(e)}")

                    if col2.button("Удалить риск", key="register_risk_delete_btn"):
                        try:
                            session.query(RiskEvaluation).filter(
                                RiskEvaluation.risk_id == selected_risk_id).delete()
                            session.delete(risk_obj)
                            session.commit()
                            st.success("Риск и связанные с ним оценки удалены!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при удалении риска: {str(e)}")
        else:
            st.info("Нет рисков. Добавьте новый риск.")
    finally:
        session.close()


def show_expert_evaluation_tab():
    """Вкладка с экспертной оценкой рисков."""
    st.subheader("Экспертная оценка рисков")
    session = SessionLocal()

    try:
        risks = session.query(RiskRegister).all()

        if not risks:
            st.info("Нет рисков для оценки. Сначала добавьте риски на вкладке 'Риски'.")
            return

        risk_options = {f"{r.code}: {r.name}": r.id for r in risks}
        selected_risk_option = st.selectbox("Выберите риск для оценки", list(risk_options.keys()),
                                            key="select_risk_eval")
        selected_risk_id = risk_options[selected_risk_option]

        risk = session.get(RiskRegister, selected_risk_id)

        if not risk:
            return

        # Информация о риске
        with st.container(border=True):
            st.markdown(f"**Риск {risk.code}: {risk.name}**")
            project_name = risk.project.name if risk.project else "Неизвестный проект"
            st.write(f"Этап: {risk.stage} | Проект: {project_name}")

        # Выбор метода расчета
        st.write("Метод расчета рейтинга:")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            basic_method = st.toggle("Базовый (умножение)", value=True, key="basic_method")
        with col2:
            weighted_method = st.toggle("Взвешенный", value=False, key="weighted_method")
        with col3:
            categorical_method = st.toggle("Категорийный", value=False, key="categorical_method")
        with col4:
            complex_method = st.toggle("Комплексный", value=False, key="complex_method")

        # Управление экспертами и их оценками
        st.write("### Экспертные оценки")

        # Секция добавления эксперта и его оценки
        with st.expander("Добавить эксперта и оценку", expanded=False):
            experts = session.query(Expert).all()

            use_existing = st.checkbox("Использовать существующего эксперта",
                                       value=True if experts else False,
                                       key="use_existing_expert")

            expert_id = None

            if use_existing and experts:
                expert_options = {f"{e.name} ({e.position})": e.id for e in experts}
                selected_expert = st.selectbox("Выберите эксперта", list(expert_options.keys()),
                                               key="select_expert")
                expert_id = expert_options[selected_expert]
            else:
                expert_name = st.text_input("ФИО эксперта", key="new_expert_name")
                expert_position = st.text_input("Должность эксперта", key="new_expert_position")

                if st.button("Добавить эксперта", key="add_expert_btn"):
                    if expert_name and expert_position:
                        new_expert = Expert(name=expert_name, position=expert_position)
                        try:
                            session.add(new_expert)
                            session.commit()
                            st.success(f"Эксперт {expert_name} добавлен!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при добавлении эксперта: {str(e)}")
                    else:
                        st.error("ФИО и должность эксперта обязательны")

            if expert_id:
                probability = st.slider("Вероятность (1-5)", 1.0, 5.0, 3.0, 1.0, key="new_eval_prob")
                impact_val = st.slider("Влияние (1-5)", 1.0, 5.0, 3.0, 1.0, key="new_eval_impact")

                current_date = datetime.datetime.now().strftime("%d.%m.%Y")

                if st.button("Сохранить оценку", key="save_eval_btn"):
                    try:
                        existing_eval = session.query(RiskEvaluation).filter(
                            RiskEvaluation.risk_id == selected_risk_id,
                            RiskEvaluation.expert_id == expert_id
                        ).first()

                        if existing_eval:
                            st.warning("Эксперт уже оценил этот риск. Обновляем оценку.")
                            existing_eval.probability = probability
                            existing_eval.impact = impact_val
                            existing_eval.evaluation_date = current_date
                            session.commit()
                            st.success("Оценка обновлена!")
                            st.rerun()
                        else:
                            evaluation = RiskEvaluation(
                                risk_id=selected_risk_id,
                                expert_id=expert_id,
                                probability=probability,
                                impact=impact_val,
                                evaluation_date=current_date
                            )
                            session.add(evaluation)
                            session.commit()
                            st.success("Оценка добавлена!")
                            st.rerun()
                    except Exception as e:
                        session.rollback()
                        st.error(f"Ошибка при сохранении оценки: {str(e)}")

        # Отображение текущих оценок для выбранного риска
        evaluations = session.query(RiskEvaluation).filter(
            RiskEvaluation.risk_id == selected_risk_id).all()

        if evaluations:
            eval_data = []
            for e in evaluations:
                expert_name = e.expert.name if e.expert else f"Эксперт ID: {e.expert_id}"
                expert_position = e.expert.position if e.expert else "Неизвестно"

                eval_data.append({
                    "Эксперт": expert_name,
                    "Должность": expert_position,
                    "Вероятность (1-5)": e.probability,
                    "Влияние (1-5)": e.impact,
                    "Дата оценки": e.evaluation_date
                })

            df_evals = pd.DataFrame(eval_data)
            st.dataframe(df_evals, hide_index=True)

            # Секция удаления оценки
            with st.expander("Удалить оценку эксперта", expanded=False):
                expert_names = [d["Эксперт"] for d in eval_data]
                if expert_names:
                    expert_to_delete = st.selectbox("Выберите эксперта для удаления оценки",
                                                    expert_names, key="delete_eval_expert")

                    if st.button("Удалить оценку", key="delete_eval_btn"):
                        try:
                            expert_id_to_delete = next(
                                (ev.expert_id for ev in evaluations if ev.expert.name == expert_to_delete),
                                None)
                            if expert_id_to_delete:
                                session.query(RiskEvaluation).filter(
                                    RiskEvaluation.risk_id == selected_risk_id,
                                    RiskEvaluation.expert_id == expert_id_to_delete
                                ).delete()
                                session.commit()
                                st.success(f"Оценка эксперта {expert_to_delete} удалена!")
                                st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка при удалении оценки: {str(e)}")

            # ИСПРАВЛЕНО: Результаты оценки вынесены из блока удаления на правильный уровень
            avg_prob = sum(e.probability for e in evaluations) / len(evaluations)
            avg_impact = sum(e.impact for e in evaluations) / len(evaluations)

            rating_basic = calculate_basic_risk(avg_prob, avg_impact)
            rating_weighted = calculate_weighted_risk(avg_prob, avg_impact)
            category, color = get_risk_category(rating_basic)
            rating_complex = calculate_complex_risk(avg_prob, avg_impact, 1.1)

            with st.container(border=True):
                st.subheader("Результаты оценки по выбранному методу:")

                col1, col2 = st.columns([1, 1])

                with col1:
                    st.write("Средняя оценка вероятности:", f"{avg_prob:.1f}")
                    st.write("Средняя оценка влияния:", f"{avg_impact:.1f}")

                    if basic_method:
                        st.write("Рейтинг риска (вероятность x влияние):", f"{rating_basic:.1f}")
                        display_risk_category(rating_basic)
                    elif weighted_method:
                        st.write("Рейтинг риска (взвешенная оценка):", f"{rating_weighted:.2f}")
                        display_risk_category(rating_weighted)
                    elif categorical_method:
                        st.write("Категорийный рейтинг:", f"{rating_basic:.1f}")
                        display_risk_category(rating_basic)
                    elif complex_method:
                        st.write("Рейтинг риска (комплексная оценка):", f"{rating_complex:.2f}")
                        display_risk_category(rating_complex)
                    else:
                        st.write("Рейтинг риска (вероятность x влияние):", f"{rating_basic:.1f}")
                        display_risk_category(rating_basic)

                with col2:
                    st.write("Матрица риска")
                    try:
                        fig = create_risk_matrix()
                        fig = mark_risk_position(fig, avg_prob, avg_impact)
                        st.plotly_chart(fig, width='stretch')
                    except Exception as e:
                        st.error(f"Ошибка при создании матрицы рисков: {str(e)}")
        else:
            st.info("Нет оценок для этого риска. Добавьте экспертные оценки.")
    finally:
        session.close()


# ==================== АКТИВЫ ====================
def show_assets_page():
    st.subheader("Управление активами")
    session = SessionLocal()

    try:
        with st.expander("Добавить актив"):
            name = st.text_input("Название актива", key="asset_add_name")
            category = st.text_input("Категория", key="asset_add_cat")
            value = st.number_input("Ценность", 0.0, 1e9, 1000.0, step=1.0, format="%.0f", key="asset_add_val")
            crit = st.slider("Критичность", 0.0, 1.0, 0.5, key="asset_add_crit")
            if st.button("Сохранить актив", key="asset_save_btn"):
                if name:
                    new_a = Asset(name=name, category=category, value=value, criticality=crit)
                    try:
                        session.add(new_a)
                        session.commit()
                        st.success("Актив добавлен!")
                        st.rerun()
                    except IntegrityError:
                        session.rollback()
                        st.error("Актив с таким названием уже существует!")
                else:
                    st.error("Название не может быть пустым")

        assets = session.query(Asset).all()
        if assets:
            df = pd.DataFrame([[a.id, a.name, a.category, a.value, a.criticality] for a in assets],
                              columns=["ID", "Name", "Category", "Value", "Criticality"])
            st.dataframe(translate_columns(df), hide_index=True)

            with st.expander("Редактировать/Удалить"):
                ids = [a.id for a in assets]
                sel_id = st.selectbox("Выберите ID актива", ids, key="asset_edit_id")
                if sel_id:
                    asset_obj = session.get(Asset, sel_id)
                    new_name = st.text_input("Новое название", asset_obj.name, key="asset_edit_name")
                    new_cat = st.text_input("Новая категория", asset_obj.category, key="asset_edit_cat")
                    new_val = st.number_input("Новая ценность", 0.0, 1e9, asset_obj.value, key="asset_edit_val")
                    new_crit = st.slider("Новая критичность", 0.0, 1.0, asset_obj.criticality, key="asset_edit_crit")

                    col1, col2 = st.columns(2)
                    if col1.button("Обновить", key="asset_update_btn"):
                        try:
                            asset_obj.name = new_name
                            asset_obj.category = new_cat
                            asset_obj.value = new_val
                            asset_obj.criticality = new_crit
                            session.commit()
                            st.success("Актив обновлён!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
                    if col2.button("Удалить", key="asset_delete_btn"):
                        try:
                            session.delete(asset_obj)
                            session.commit()
                            st.success("Актив удалён!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
        else:
            st.info("Пока нет активов.")
    finally:
        session.close()


# ==================== УГРОЗЫ ====================
def show_threats_page():
    st.subheader("Управление угрозами")
    session = SessionLocal()

    try:
        with st.expander("Добавить угрозу"):
            name = st.text_input("Название угрозы", key="threat_add_name")
            ttype = st.text_input("Тип угрозы", key="threat_add_type")
            prob = st.slider("Вероятность", 0.0, 1.0, 0.5, key="threat_add_prob")
            desc = st.text_area("Описание", key="threat_add_desc")
            if st.button("Сохранить угрозу", key="threat_save_btn"):
                if name:
                    th = Threat(name=name, type=ttype, probability=prob, description=desc)
                    try:
                        session.add(th)
                        session.commit()
                        st.success("Угроза добавлена!")
                        st.rerun()
                    except IntegrityError:
                        session.rollback()
                        st.error("Угроза с таким названием уже существует!")
                else:
                    st.error("Название не может быть пустым")

        threats = session.query(Threat).all()
        if threats:
            df = pd.DataFrame([[t.id, t.name, t.type, t.probability, t.description] for t in threats],
                              columns=["ID", "Name", "Type", "Probability", "Description"])
            st.dataframe(translate_columns(df), hide_index=True)

            with st.expander("Редактировать/Удалить"):
                ids = [t.id for t in threats]
                sel_id = st.selectbox("Выберите ID угрозы", ids, key="threat_edit_id")
                if sel_id:
                    threat_obj = session.get(Threat, sel_id)
                    new_name = st.text_input("Новое название", threat_obj.name, key="threat_edit_name")
                    new_type = st.text_input("Новый тип", threat_obj.type, key="threat_edit_type")
                    new_prob = st.slider("Новая вероятность", 0.0, 1.0, threat_obj.probability, key="threat_edit_prob")
                    new_desc = st.text_area("Новое описание", threat_obj.description, key="threat_edit_desc")
                    col1, col2 = st.columns(2)
                    if col1.button("Обновить", key="threat_update_btn"):
                        try:
                            threat_obj.name = new_name
                            threat_obj.type = new_type
                            threat_obj.probability = new_prob
                            threat_obj.description = new_desc
                            session.commit()
                            st.success("Угроза обновлена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
                    if col2.button("Удалить", key="threat_delete_btn"):
                        try:
                            session.delete(threat_obj)
                            session.commit()
                            st.success("Угроза удалена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
        else:
            st.info("Пока нет угроз.")
    finally:
        session.close()


# ==================== УЯЗВИМОСТИ ====================
def show_vulnerabilities_page():
    st.subheader("Управление уязвимостями")
    session = SessionLocal()

    try:
        with st.expander("Добавить уязвимость"):
            name = st.text_input("Название", key="vuln_add_name")
            level = st.slider("Уровень", 0.0, 1.0, 0.5, key="vuln_add_level")
            desc = st.text_area("Описание", key="vuln_add_desc")
            if st.button("Сохранить уязвимость", key="vuln_save_btn"):
                if name:
                    v = Vulnerability(name=name, level=level, description=desc)
                    try:
                        session.add(v)
                        session.commit()
                        st.success("Уязвимость добавлена!")
                        st.rerun()
                    except IntegrityError:
                        session.rollback()
                        st.error("Уязвимость с таким названием уже существует!")
                else:
                    st.error("Название не может быть пустым")

        vulns = session.query(Vulnerability).all()
        if vulns:
            df = pd.DataFrame([[v.id, v.name, v.level, v.description] for v in vulns],
                              columns=["ID", "Name", "Level", "Description"])
            st.dataframe(translate_columns(df), hide_index=True)

            with st.expander("Редактировать/Удалить"):
                ids = [v.id for v in vulns]
                sel_id = st.selectbox("Выберите ID уязвимости", ids, key="vuln_edit_id")
                if sel_id:
                    v_obj = session.get(Vulnerability, sel_id)
                    new_name = st.text_input("Новое название", v_obj.name, key="vuln_edit_name")
                    new_level = st.slider("Новый уровень", 0.0, 1.0, v_obj.level, key="vuln_edit_level")
                    new_desc = st.text_area("Новое описание", v_obj.description, key="vuln_edit_desc")
                    col1, col2 = st.columns(2)
                    if col1.button("Обновить", key="vuln_update_btn"):
                        try:
                            v_obj.name = new_name
                            v_obj.level = new_level
                            v_obj.description = new_desc
                            session.commit()
                            st.success("Уязвимость обновлена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
                    if col2.button("Удалить", key="vuln_delete_btn"):
                        try:
                            session.delete(v_obj)
                            session.commit()
                            st.success("Уязвимость удалена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
        else:
            st.info("Пока нет уязвимостей.")
    finally:
        session.close()


# ==================== КОНТРОЛИ ====================
def show_controls_page():
    st.subheader("Управление мерами контроля")
    session = SessionLocal()

    try:
        with st.expander("Добавить меру"):
            cname = st.text_input("Название меры", key="ctrl_add_name")
            ctype = st.text_input("Тип меры", key="ctrl_add_type")
            eff = st.slider("Эффективность", 0.0, 1.0, 0.5, key="ctrl_add_eff")
            cost = st.number_input("Стоимость", 0.0, 100000000.0, 0.0, step=1.0, format="%.0f", key="ctrl_add_cost")
            desc = st.text_area("Описание", key="ctrl_add_desc")
            if st.button("Сохранить меру", key="ctrl_save_btn"):
                if cname:
                    c = Control(name=cname, type=ctype, effectiveness=eff, cost=cost, description=desc)
                    try:
                        session.add(c)
                        session.commit()
                        st.success("Мера контроля добавлена!")
                        st.rerun()
                    except IntegrityError:
                        session.rollback()
                        st.error("Мера с таким названием уже существует!")
                else:
                    st.error("Название не может быть пустым")

        controls = session.query(Control).all()
        if controls:
            df = pd.DataFrame([[c.id, c.name, c.type, c.effectiveness, c.cost, c.description]
                               for c in controls],
                              columns=["ID", "Name", "Type", "Effectiveness", "Cost", "Description"])
            st.dataframe(translate_columns(df), hide_index=True)

            with st.expander("Редактировать/Удалить"):
                ids = [c.id for c in controls]
                sel_id = st.selectbox("Выберите ID меры", ids, key="ctrl_edit_id")
                if sel_id:
                    c_obj = session.get(Control, sel_id)
                    new_name = st.text_input("Новое название", c_obj.name, key="ctrl_edit_name")
                    new_type = st.text_input("Новый тип", c_obj.type, key="ctrl_edit_type")
                    new_eff = st.slider("Новая эффективность", 0.0, 1.0, c_obj.effectiveness, key="ctrl_edit_eff")
                    new_cost = st.number_input("Новая стоимость", 0.0, 1e6, c_obj.cost, step=1.0, format="%.0f",
                                               key="ctrl_edit_cost")
                    new_desc = st.text_area("Новое описание", c_obj.description, key="ctrl_edit_desc")
                    col1, col2 = st.columns(2)
                    if col1.button("Обновить", key="ctrl_update_btn"):
                        try:
                            c_obj.name = new_name
                            c_obj.type = new_type
                            c_obj.effectiveness = new_eff
                            c_obj.cost = new_cost
                            c_obj.description = new_desc
                            session.commit()
                            st.success("Мера обновлена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
                    if col2.button("Удалить", key="ctrl_delete_btn"):
                        try:
                            session.delete(c_obj)
                            session.commit()
                            st.success("Мера удалена!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
        else:
            st.info("Пока нет мер контроля.")
    finally:
        session.close()


# ==================== СЦЕНАРИИ РИСКА ====================
def show_risk_scenarios_page():
    st.subheader("Сценарии риска")
    session = SessionLocal()

    try:
        assets = session.query(Asset).all()
        threats = session.query(Threat).all()
        vulns = session.query(Vulnerability).all()
        controls = session.query(Control).all()

        if assets and threats and vulns:
            with st.expander("Создать сценарий риска"):
                asset_opt = {f"{a.name} (ID:{a.id})": a.id for a in assets}
                threat_opt = {f"{t.name} (ID:{t.id})": t.id for t in threats}
                vuln_opt = {f"{v.name} (ID:{v.id})": v.id for v in vulns}
                ctrl_opt = {f"{c.name} (ID:{c.id})": c.id for c in controls}

                sa = st.selectbox("Актив", list(asset_opt.keys()), key="scenario_asset")
                stt = st.selectbox("Угроза", list(threat_opt.keys()), key="scenario_threat")
                sv = st.selectbox("Уязвимость", list(vuln_opt.keys()), key="scenario_vuln")
                sc = st.selectbox("Контроль (необязательно)", ["Нет"] + list(ctrl_opt.keys()),
                                  key="scenario_ctrl")

                if st.button("Сохранить сценарий", key="scenario_save_btn"):
                    ctrl_id = None
                    if sc != "Нет":
                        ctrl_id = ctrl_opt[sc]
                    scenario = RiskScenario(
                        asset_id=asset_opt[sa],
                        threat_id=threat_opt[stt],
                        vuln_id=vuln_opt[sv],
                        control_id=ctrl_id
                    )
                    try:
                        session.add(scenario)
                        session.commit()
                        st.success("Сценарий риска создан!")
                        st.rerun()
                    except Exception as e:
                        session.rollback()
                        st.error(f"Ошибка: {str(e)}")
        else:
            st.warning("Нужно иметь хотя бы один актив, угрозу и уязвимость.")

        scenarios = session.query(RiskScenario).all()
        if scenarios:
            rows = []
            for s in scenarios:
                base_r, mitig_r = calculate_risk(
                    s.asset.value,
                    s.threat.probability,
                    s.vuln.level,
                    s.control.effectiveness if s.control else 0
                )
                rows.append({
                    "ID": s.id,
                    "Asset": s.asset.name,
                    "Threat": s.threat.name,
                    "Vulnerability": s.vuln.name,
                    "Control": s.control.name if s.control else "Нет",
                    "Base Risk": base_r,
                    "Mitigated Risk": mitig_r
                })
            df = pd.DataFrame(rows)
            st.dataframe(translate_columns(df), hide_index=True)

            with st.expander("Редактировать/Удалить сценарий"):
                ids = [s.id for s in scenarios]
                sel_id = st.selectbox("Выберите ID сценария", ids, key="scenario_edit_id")
                if sel_id:
                    scen_obj = session.get(RiskScenario, sel_id)
                    st.write(f"Текущий контроль: {scen_obj.control.name if scen_obj.control else 'Нет'}")

                    ctrl_choices_map = {"Убрать контроль": None}
                    for c in controls:
                        ctrl_choices_map[f"{c.name} (ID:{c.id})"] = c.id

                    chosen_ctrl = st.selectbox("Контроль", list(ctrl_choices_map.keys()),
                                               key="scenario_edit_ctrl")

                    col1, col2 = st.columns(2)
                    if col1.button("Обновить", key="scenario_update_btn"):
                        try:
                            scen_obj.control_id = ctrl_choices_map[chosen_ctrl]
                            session.commit()
                            st.success("Сценарий обновлён!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
                    if col2.button("Удалить", key="scenario_delete_btn"):
                        try:
                            session.delete(scen_obj)
                            session.commit()
                            st.success("Сценарий удалён!")
                            st.rerun()
                        except Exception as e:
                            session.rollback()
                            st.error(f"Ошибка: {str(e)}")
        else:
            st.info("Пока нет сценариев риска.")
    finally:
        session.close()


# ==================== АНАЛИТИКА ====================
def show_analytics_page():
    st.subheader("Аналитика и отчёты")
    session = SessionLocal()

    try:
        scenarios = session.query(RiskScenario).all()
        if scenarios:
            data = []
            total_base = 0
            total_mitig = 0
            for s in scenarios:
                base_r, mitig_r = calculate_risk(
                    s.asset.value,
                    s.threat.probability,
                    s.vuln.level,
                    s.control.effectiveness if s.control else 0
                )
                data.append({
                    "ScenarioID": s.id,
                    "Asset": s.asset.name,
                    "Threat": s.threat.name,
                    "Vulnerability": s.vuln.name,
                    "Control": s.control.name if s.control else "Нет",
                    "Base Risk": base_r,
                    "Mitigated Risk": mitig_r
                })
                total_base += base_r
                total_mitig += mitig_r

            df_report = pd.DataFrame(data)
            st.dataframe(translate_columns(df_report))
            st.metric("Общий базовый риск", f"{total_base:.2f}")
            st.metric("Общий остаточный риск", f"{total_mitig:.2f}")

            st.write("### Тепловая карта базового риска (Актив vs Угроза)")
            pivot_base = df_report.pivot_table(
                values="Base Risk",
                index="Threat",
                columns="Asset",
                aggfunc="mean",
                fill_value=0
            )
            fig = px.imshow(
                pivot_base,
                color_continuous_scale="Reds",
                title="Тепловая карта",
                labels={"x": "Активы", "y": "Угрозы", "color": "Базовый риск"}
            )
            st.plotly_chart(style_fig(fig), width='stretch')

            # Экспорт в Excel — один клик
            st.write("### Выгрузить отчёт в Excel")
            xlsx_data = export_to_excel({"Risk Report": df_report})
            st.download_button(
                label="Скачать risk_report.xlsx",
                data=xlsx_data,
                file_name="risk_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.info("Нет данных для аналитики.")
    finally:
        session.close()


# ==================== 3D ВИЗУАЛИЗАЦИЯ ====================
def show_3d_visualization_page():
    st.subheader("3D-график: Актив-Угроза-Уязвимость")
    session = SessionLocal()

    try:
        scenarios = session.query(RiskScenario).all()
        if not scenarios:
            st.info("Нет сценариев риска для 3D-графа.")
            return

        G = nx.Graph()
        for s in scenarios:
            if s.asset and s.threat and s.vuln:
                asset_node = ("asset", s.asset.id, s.asset.name)
                threat_node = ("threat", s.threat.id, s.threat.name)
                vuln_node = ("vuln", s.vuln.id, s.vuln.name)

                G.add_node(asset_node)
                G.add_node(threat_node)
                G.add_node(vuln_node)

                G.add_edge(asset_node, threat_node)
                G.add_edge(threat_node, vuln_node)
                G.add_edge(asset_node, vuln_node)

        if G.number_of_nodes() == 0:
            st.info("Не получилось собрать узлы. Проверьте, что у сценариев есть актив/угроза/уязвимость.")
            return

        pos3d = nx.spring_layout(G, dim=3, k=0.5, seed=42)

        x_edges, y_edges, z_edges = [], [], []
        for edge in G.edges():
            x0, y0, z0 = pos3d[edge[0]]
            x1, y1, z1 = pos3d[edge[1]]
            x_edges += [x0, x1, None]
            y_edges += [y0, y1, None]
            z_edges += [z0, z1, None]

        x_nodes, y_nodes, z_nodes, text_nodes = [], [], [], []
        for node in G.nodes():
            x_nodes.append(pos3d[node][0])
            y_nodes.append(pos3d[node][1])
            z_nodes.append(pos3d[node][2])
            text_nodes.append(f"{node[0]} #{node[1]} - {node[2]}")

        fig = go.Figure()
        fig.add_trace(
            go.Scatter3d(
                x=x_edges, y=y_edges, z=z_edges,
                mode='lines',
                line=dict(color='gray', width=2),
                hoverinfo='none'
            )
        )
        fig.add_trace(
            go.Scatter3d(
                x=x_nodes, y=y_nodes, z=z_nodes,
                mode='markers+text',
                text=text_nodes,
                textposition="top center",
                marker=dict(size=5, color='cornflowerblue'),
                hoverinfo='text'
            )
        )
        fig.update_layout(
            showlegend=False,
            scene=dict(
                xaxis=dict(showbackground=True),
                yaxis=dict(showbackground=True),
                zaxis=dict(showbackground=True)
            )
        )
        st.plotly_chart(fig, width='stretch')
    finally:
        session.close()


# ==================== ЭКОНОМИЧЕСКАЯ ОЦЕНКА ====================
def show_economic_page():
    st.subheader("Экономическая оценка (EVI, ROI, TCO)")
    st.markdown(
        '<span class="rn-pill">EVI — основной индекс методики</span>',
        unsafe_allow_html=True,
    )
    st.caption(
        "В методике диссертации экономическая приоритизация мер защиты выполняется через индекс EVI "
        "(снижение риска на единицу затрат). Показатели ROI и TCO приведены как вспомогательные."
    )

    tab_evi, tab_roi, tab_tco = st.tabs(["EVI — приоритизация мер", "ROI", "TCO"])

    # ---------- EVI ----------
    with tab_evi:
        st.markdown("##### Ручной расчёт EVI")
        col1, col2 = st.columns(2)
        with col1:
            delta_r = st.number_input("Ожидаемое снижение риска ΔR (₽)", 0.0, 1e12, 100000.0,
                                      step=1000.0, format="%.0f", key="evi_delta")
        with col2:
            cost_m = st.number_input("Совокупные затраты на меру (₽)", 0.0, 1e12, 60000.0,
                                     step=1000.0, format="%.0f", key="evi_cost")
        if st.button("Рассчитать EVI", key="evi_calc_btn", type="primary"):
            evi = calculate_evi(delta_r, cost_m)
            if evi is None:
                st.error("Затраты должны быть больше 0.")
            else:
                verdict = "мера окупается по риску" if evi >= 1 else "снижение риска меньше затрат"
                st.success(f"EVI = {evi:.2f}  ({verdict})")

        st.divider()
        st.markdown("##### Приоритизация мер по сценариям риска")
        session = SessionLocal()
        try:
            scenarios = [s for s in session.query(RiskScenario).all()
                         if s.asset and s.threat and s.vuln and s.control]
            if not scenarios:
                st.info("Нет сценариев с назначенными мерами контроля. "
                        "Создайте сценарии и привяжите к ним контроли — таблица приоритизации появится здесь.")
            else:
                rows = []
                for s in scenarios:
                    base_r, mitig_r = calculate_risk(s.asset.value, s.threat.probability,
                                                     s.vuln.level, s.control.effectiveness)
                    delta = base_r - mitig_r
                    evi = calculate_evi(delta, s.control.cost)
                    rows.append({
                        "Сценарий": f"{s.asset.name} ← {s.threat.name}",
                        "Мера": s.control.name,
                        "Снижение риска ΔR (₽)": round(delta),
                        "Затраты (₽)": round(s.control.cost or 0),
                        "EVI": round(evi, 2) if evi is not None else None,
                    })
                df = pd.DataFrame(rows).sort_values("EVI", ascending=False, na_position="last")
                st.dataframe(df, hide_index=True, width='stretch')

                plot_df = df.dropna(subset=["EVI"])
                if not plot_df.empty:
                    fig = px.bar(plot_df, x="EVI", y="Мера", orientation="h",
                                 color="EVI", color_continuous_scale=["#DC2626", "#F59E0B", "#16A34A"],
                                 title="Приоритет мер защиты по EVI")
                    fig.add_vline(x=1.0, line_dash="dash", line_color=MUTED,
                                  annotation_text="EVI = 1", annotation_position="top")
                    fig.update_layout(coloraxis_showscale=False, yaxis=dict(autorange="reversed"))
                    st.plotly_chart(style_fig(fig, height=80 + 46 * len(plot_df)), width='stretch')
                    st.caption("Меры с EVI ≥ 1 окупаются по снижению риска; чем выше EVI, тем выше приоритет внедрения.")
        finally:
            session.close()

    # ---------- ROI ----------
    with tab_roi:
        st.caption("Вспомогательный показатель. В выводы диссертации ROI не вводится.")
        cost = st.number_input("Стоимость (Cost)", 0.0, 1e9, 1000.0, format="%.0f", step=1.0, key="econ_cost")
        benefit = st.number_input("Выгода (Benefit)", 0.0, 1e9, 1500.0, format="%.0f", step=1.0, key="econ_benefit")
        if st.button("Рассчитать ROI", key="econ_roi_btn"):
            roi = calculate_roi(cost, benefit)
            if roi is None:
                st.error("Cost должен быть > 0.")
            else:
                st.success(f"ROI: {roi:.2f}%")

    # ---------- TCO ----------
    with tab_tco:
        init_cost = st.number_input("Начальная стоимость", 0.0, 1e9, 5000.0, step=1.0, format="%.0f", key="econ_init")
        op_cost = st.number_input("Операционные расходы (в год)", 0.0, 1e9, 1000.0, step=1.0, format="%.0f", key="econ_op")
        years = st.number_input("Срок (лет)", 1.0, 50.0, 5.0, format="%.0f", step=1.0, key="econ_years")
        if st.button("Рассчитать TCO", key="econ_tco_btn"):
            tco = calculate_tco(init_cost, op_cost, years)
            st.success(f"Итоговый TCO: {tco:,.2f} ₽".replace(",", " "))


# ==================== ДИНАМИЧЕСКИЕ ДАШБОРДЫ ====================
def show_dashboards_page():
    st.subheader("Динамические дашборды: ключевые метрики управления рисками")
    session = SessionLocal()

    try:
        scenarios = session.query(RiskScenario).all()

        if not scenarios:
            st.info("Нет сценариев для дашборда. Создайте сценарии риска.")
            return

        assets = session.query(Asset).all()
        threats = session.query(Threat).all()
        vulns = session.query(Vulnerability).all()
        controls = session.query(Control).all()

        # Подготовка метрик
        total_assets_value = sum(a.value for a in assets if a.value is not None)
        total_threats = len(threats)
        total_vulnerabilities = len(vulns)
        total_controls = len(controls)

        # Подсчёт рисков через ORM (без N+1)
        total_base_risk = 0
        total_mitigated_risk = 0
        risk_data = []

        for s in scenarios:
            if s.asset and s.threat and s.vuln:
                base_risk, mitigated_risk = calculate_risk(
                    s.asset.value,
                    s.threat.probability,
                    s.vuln.level,
                    s.control.effectiveness if s.control else 0
                )
                total_base_risk += base_risk
                total_mitigated_risk += mitigated_risk
                risk_data.append({
                    "Актив": s.asset.name,
                    "Базовый риск": base_risk,
                    "Сниженный риск": mitigated_risk
                })

        # Ключевые метрики
        st.markdown("### Ключевые метрики")
        col1, col2, col3 = st.columns(3)
        col1.metric("Суммарная ценность активов", f"{total_assets_value:,.2f}")
        col2.metric("Общее количество угроз", total_threats)
        col3.metric("Общее количество уязвимостей", total_vulnerabilities)

        col1, col2 = st.columns(2)
        col1.metric("Общий базовый риск", f"{total_base_risk:,.2f}")
        col2.metric("Общий остаточный риск", f"{total_mitigated_risk:,.2f}")

        st.metric("Общее количество мер контроля", total_controls)

        # Визуализация риска по активам
        if risk_data:
            st.markdown("### Риск по активам")
            df_risk = pd.DataFrame(risk_data)

            # Группируем по активу (на случай нескольких сценариев для одного актива)
            df_grouped = df_risk.groupby("Актив").sum().reset_index()

            fig = px.bar(
                df_grouped,
                x="Актив",
                y=["Базовый риск", "Сниженный риск"],
                barmode="group",
                title="Базовый и сниженный риск по активам",
                labels={"value": "Риск", "variable": "Тип риска"}
            )
            st.plotly_chart(style_fig(fig), width='stretch')
    finally:
        session.close()


# ==================== СПРАВКА ====================
def show_reference_page():
    st.subheader("Методики оценки вероятности и уязвимости")
    st.markdown("""
    ### Определение вероятности угрозы
    - Высокая (близко к 1): Частые инциденты, активная эксплуатация уязвимостей.
    - Средняя (около 0.5): Возможны сценарии, но они редки.
    - Низкая (0-0.3): Угроза маловероятна из-за отсутствия интереса или сложностей.

    ### Определение уровня уязвимости
    - Высокая (близко к 1): Серьезные уязвимости без мер защиты.
    - Средняя (0.5): Есть уязвимости, но частично компенсируются контролями.
    - Низкая (0-0.3): Хорошо защищенные системы, обновленные и протестированные.

    ### Рекомендации
    - Используйте базу рисков для уточнения вероятностей.
    - Ссылайтесь на стандарты (NIST, ISO), публичные базы уязвимостей (NVD) для обоснования оценок.
    - Периодически пересматривайте оценки по мере изменения ландшафта угроз.
    """)


# ==================== КАЛЬКУЛЯТОРЫ ====================
def show_calculators_page():
    st.subheader("Калькуляторы риска")
    st.markdown("#### Простой калькулятор")
    asset_value = st.number_input("Стоимость актива", min_value=0.0, value=1000.0, step=1.0, format="%.0f",
                                  key="calc_asset")
    threat_prob = st.slider("Вероятность угрозы (0-1)", 0.0, 1.0, 0.5, key="calc_threat")
    vuln_level = st.slider("Уровень уязвимости (0-1)", 0.0, 1.0, 0.5, key="calc_vuln")
    control_eff = st.slider("Эффективность контроля (0-1)", 0.0, 1.0, 0.0, key="calc_ctrl")
    if st.button("Рассчитать риск", key="calc_btn"):
        base_risk, mitigated_risk = calculate_risk(asset_value, threat_prob, vuln_level, control_eff)
        st.write(f"Базовый риск: {base_risk:.2f}")
        st.write(f"С учетом контроля: {mitigated_risk:.2f}")


# ==================== ОЦЕНКА ЗРЕЛОСТИ ====================
def show_maturity_page():
    st.subheader("Оценка зрелости процессов управления рисками ИБ")
    st.markdown(
        '<div class="rn-card">Модуль реализует модель факторов результативности (15 факторов в 5 категориях, '
        'раздел 3.4) и адаптированную модель зрелости с формулой прогнозирования (приложение В.1). '
        'Результат — профиль зрелости и приоритетные направления улучшения.</div>',
        unsafe_allow_html=True,
    )

    tab_q, tab_ml, tab_hist = st.tabs(["Оценка по 15 факторам", "Прогноз по формуле (ML)", "История оценок"])

    # ---------- Анкета по 15 факторам ----------
    with tab_q:
        st.caption("Оцените каждый фактор по шкале 1–5 (1 — отсутствует, 5 — полностью реализован).")

        with st.expander("Веса факторов (по умолчанию — по диссертации, можно изменить)"):
            st.caption(
                "Веса для факторов «простота методологии» (9,5), «поддержка руководства» (9,3), "
                "«фокус на критических рисках» (9,2) и «компетенции сотрудников» (9,1) соответствуют "
                "экспертным оценкам влияния из диссертации; остальные заданы нейтрально (8,0) и подлежат калибровке."
            )
            weights = {}
            for cat, factors in MATURITY_FACTORS.items():
                cols = st.columns(len(factors))
                for col, (fname, fweight) in zip(cols, factors):
                    weights[fname] = col.number_input(
                        fname, 1.0, 10.0, float(fweight), step=0.1,
                        key=f"w_{fname}", format="%.1f",
                    )

        scores = {}
        for cat, factors in MATURITY_FACTORS.items():
            st.markdown(f"**{cat}**")
            for fname, _ in factors:
                scores[fname] = st.slider(fname, 1, 5, 3, key=f"mf_{fname}")

        # Средневзвешенная оценка (1..5)
        wsum = sum(weights.get(f, 8.0) for f in scores)
        overall = sum(scores[f] * weights.get(f, 8.0) for f in scores) / wsum if wsum else 0
        band_name, band_desc = maturity_band(overall)

        st.divider()
        m1, m2 = st.columns([1, 2])
        with m1:
            st.metric("Общий уровень зрелости", f"{overall:.2f} / 5")
            st.markdown(f'<span class="rn-pill">{band_name}</span>', unsafe_allow_html=True)
            st.caption(band_desc)
        with m2:
            radar = go.Figure()
            radar.add_trace(go.Scatterpolar(
                r=[scores[f] for f in scores] + [scores[list(scores)[0]]],
                theta=list(scores.keys()) + [list(scores.keys())[0]],
                fill="toself", line=dict(color=ACCENT), fillcolor="rgba(47,111,237,0.18)",
                name="Профиль",
            ))
            radar.update_layout(
                polar=dict(radialaxis=dict(range=[0, 5], tickvals=[1, 2, 3, 4, 5], gridcolor="#E3E8F0"),
                           angularaxis=dict(gridcolor="#E3E8F0")),
                showlegend=False, height=440,
                font=dict(family=PLOTLY_FONT, color=INK, size=11),
                paper_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=80, r=80, t=40, b=40),
            )
            st.plotly_chart(radar, width='stretch')

        # Слабые места
        weakest = sorted(scores.items(), key=lambda kv: (kv[1], -weights.get(kv[0], 8.0)))[:3]
        st.markdown("##### Приоритетные направления улучшения")
        st.markdown("\n".join(
            f"- **{f}** — текущая оценка {v}/5 (вес {weights.get(f, 8.0):.1f})" for f, v in weakest
        ))

        st.divider()
        with st.expander("Сохранить результат оценки"):
            org = st.text_input("Название предприятия / подразделения", key="mat_org")
            comment = st.text_area("Комментарий", key="mat_comment")
            if st.button("Сохранить оценку", key="mat_save", type="primary"):
                session = SessionLocal()
                try:
                    rec = MaturityAssessment(
                        org_name=org or "—",
                        date=datetime.datetime.now().strftime("%d.%m.%Y %H:%M"),
                        overall_score=round(overall, 2),
                        ml_predicted=None,
                        factor_scores=json.dumps(scores, ensure_ascii=False),
                        comment=comment,
                    )
                    session.add(rec)
                    session.commit()
                    st.success("Оценка сохранена. См. вкладку «История оценок».")
                except Exception as e:
                    session.rollback()
                    st.error(f"Ошибка сохранения: {e}")
                finally:
                    session.close()

    # ---------- Прогноз по формуле модели зрелости ----------
    with tab_ml:
        st.markdown("##### Прогноз уровня зрелости по характеристикам предприятия")
        st.latex(r"ML = 0{,}65\cdot S + 0{,}50\cdot A + 1{,}00\cdot B + 0{,}85\cdot IT + 0{,}50\cdot IND + 0{,}50\cdot D")
        st.caption("Все показатели нормируются в диапазоне 0–1; итоговый ML — в диапазоне 0–4 (приложение В.1).")

        c1, c2, c3 = st.columns(3)
        S = c1.slider("S — размер предприятия", 0.0, 1.0, 0.4, 0.05,
                      help="0 — микро, 1 — близко к верхней границе малого предприятия", key="ml_s")
        A = c2.slider("A — возраст бизнеса", 0.0, 1.0, 0.5, 0.05, key="ml_a")
        B = c3.slider("B — доля бюджета на ИБ", 0.0, 1.0, 0.3, 0.05, key="ml_b")
        c4, c5, c6 = st.columns(3)
        IT = c4.select_slider("IT — штатный ИТ-специалист", options=[0.0, 1.0],
                              value=0.0, format_func=lambda v: "Нет" if v == 0.0 else "Есть", key="ml_it")
        IND = c5.slider("IND — отраслевой коэффициент", 0.0, 1.0, 0.5, 0.05, key="ml_ind")
        D = c6.slider("D — уровень цифровизации", 0.0, 1.0, 0.5, 0.05, key="ml_d")

        ml = 0.65 * S + 0.50 * A + 1.00 * B + 0.85 * IT + 0.50 * IND + 0.50 * D
        band_name, band_desc = maturity_band(ml)
        st.divider()
        cc1, cc2 = st.columns([1, 2])
        cc1.metric("Прогноз ML", f"{ml:.2f} / 4")
        cc1.markdown(f'<span class="rn-pill">{band_name}</span>', unsafe_allow_html=True)
        # шкала-индикатор
        gauge = go.Figure(go.Indicator(
            mode="gauge+number", value=ml, number={"font": {"size": 30}},
            gauge={
                "axis": {"range": [0, 4]},
                "bar": {"color": ACCENT},
                "steps": [
                    {"range": [0, 1], "color": "#FEE2E2"},
                    {"range": [1, 2], "color": "#FEF3C7"},
                    {"range": [2, 3], "color": "#DBEAFE"},
                    {"range": [3, 4], "color": "#DCFCE7"},
                ],
            },
        ))
        gauge.update_layout(height=240, margin=dict(l=20, r=20, t=10, b=10),
                            font=dict(family=PLOTLY_FONT, color=INK), paper_bgcolor="rgba(0,0,0,0)")
        cc2.plotly_chart(gauge, width='stretch')
        st.caption(band_desc)

    # ---------- История ----------
    with tab_hist:
        session = SessionLocal()
        try:
            records = session.query(MaturityAssessment).order_by(MaturityAssessment.id.desc()).all()
            if not records:
                st.info("Сохранённых оценок пока нет.")
            else:
                df = pd.DataFrame([{
                    "Дата": r.date, "Предприятие": r.org_name,
                    "Уровень (1–5)": r.overall_score, "Комментарий": r.comment or "",
                } for r in records])
                st.dataframe(df, hide_index=True, width='stretch')
                ids = [r.id for r in records]
                del_id = st.selectbox("ID для удаления", ids, key="mat_del_id")
                if st.button("Удалить запись", key="mat_del_btn"):
                    obj = session.get(MaturityAssessment, del_id)
                    if obj:
                        session.delete(obj)
                        session.commit()
                        st.success("Запись удалена.")
                        st.rerun()
        finally:
            session.close()


# ==================== МАСТЕР ВЫБОРА МОДЕЛИ ====================
def show_classification_wizard_page():
    st.subheader("Мастер выбора модели управления рисками ИБ")
    st.markdown(
        '<div class="rn-card">Мастер реализует классификационную матрицу из шести классов моделей '
        '(приложение Б). По характеристикам предприятия подбирается наиболее подходящий класс и профиль '
        'конфигурации программного комплекса.</div>',
        unsafe_allow_html=True,
    )

    size_opts = ["Микро (до 15 чел.)", "Малое (15–50 чел.)", "Среднее (50–100 чел.)"]
    size_keys = ["микро", "малое", "среднее"]
    mat_opts = ["Низкий", "Средний", "Высокий"]
    mat_keys = ["низкий", "средний", "высокий"]
    reg_opts = ["Низкое", "Среднее (ФЗ-152)", "Высокое (финансы, медицина, КИИ)"]
    reg_keys = ["низкая", "средняя", "высокая"]
    bud_opts = ["Ограниченный", "Средний", "Значительный"]
    bud_keys = ["ограниченный", "средний", "значительный"]
    exp_opts = ["Нет", "Частично (аутсорс)", "Штатный специалист"]
    exp_keys = ["нет", "частично", "штатный"]
    crit_opts = ["Низкая", "Средняя", "Высокая"]
    crit_keys = ["низкая", "средняя", "высокая"]

    c1, c2, c3 = st.columns(3)
    size = size_keys[size_opts.index(c1.selectbox("Размер предприятия", size_opts, key="wz_size"))]
    maturity = mat_keys[mat_opts.index(c2.selectbox("Уровень зрелости ИБ", mat_opts, index=1, key="wz_mat"))]
    regulation = reg_keys[reg_opts.index(c3.selectbox("Регулирование отрасли", reg_opts, key="wz_reg"))]
    c4, c5, c6 = st.columns(3)
    budget = bud_keys[bud_opts.index(c4.selectbox("Бюджет на ИБ", bud_opts, key="wz_bud"))]
    expertise = exp_keys[exp_opts.index(c5.selectbox("Экспертиза по ИБ", exp_opts, key="wz_exp"))]
    criticality = crit_keys[crit_opts.index(c6.selectbox("Критичность активов", crit_opts, index=1, key="wz_crit"))]

    if st.button("Подобрать модель", key="wz_btn", type="primary"):
        results = []
        for cls in MODEL_CLASSES:
            score = (cls["base"] + cls["size"][size] + cls["maturity"][maturity]
                     + cls["regulation"][regulation] + cls["budget"][budget]
                     + cls["expertise"][expertise] + cls["criticality"][criticality])
            results.append((score, cls))
        results.sort(key=lambda x: x[0], reverse=True)
        best_score, best = results[0]

        st.divider()
        st.markdown(
            f'<div class="rn-card"><span class="rn-pill">Рекомендация</span>'
            f'<h3 style="margin:8px 0 4px 0;">Класс {best["code"]}. {best["name"]}</h3>'
            f'<div style="color:{MUTED};">{best["fit"]}</div></div>',
            unsafe_allow_html=True,
        )
        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Ресурсоёмкость", best["resource"])
        p2.metric("Масштабируемость", best["scalability"])
        p3.metric("Формализация", best["formalization"])
        p4.metric("Соответствие нормам РФ", best["ru_compliance"])
        st.markdown(f"**Актуальные примеры:** {best['examples']}")

        # Ранжирование всех классов
        st.markdown("##### Пригодность всех классов для вашего профиля")
        rank_df = pd.DataFrame([{
            "Класс": f"{cls['code']}. {cls['name']}",
            "Оценка соответствия": sc,
        } for sc, cls in results])
        fig = px.bar(rank_df, x="Оценка соответствия", y="Класс", orientation="h",
                     color="Оценка соответствия", color_continuous_scale=["#CBD5E1", ACCENT])
        fig.update_layout(coloraxis_showscale=False, yaxis=dict(autorange="reversed"))
        st.plotly_chart(style_fig(fig, height=300), width='stretch')
        st.caption("Оценка отражает относительное соответствие класса введённому профилю предприятия "
                   "и является ориентиром; классы не являются взаимоисключающими и могут комбинироваться.")

    with st.expander("Справка: классификационная матрица (6 классов)"):
        ref = pd.DataFrame([{
            "Класс": f"{c['code']}. {c['name']}",
            "Ресурсоёмкость": c["resource"],
            "Масштабируемость": c["scalability"],
            "Формализация": c["formalization"],
            "Нормы РФ": c["ru_compliance"],
            "Пригодность для МСП": c["fit"],
        } for c in MODEL_CLASSES])
        st.dataframe(ref, hide_index=True, width='stretch')


# -----------------------------------------------------------
# 10. Навигация
# -----------------------------------------------------------
NAV_GROUPS = [
    ("Обзор", ["Главная"]),
    ("Оценка рисков", ["Активы", "Угрозы", "Уязвимости", "Контроли",
                       "Сценарии риска", "Реестр рисков", "Риски"]),
    ("Методика", ["Мастер выбора модели", "Оценка зрелости", "Экономика (EVI, ROI, TCO)"]),
    ("Аналитика", ["Аналитика", "Аналитика 3D", "Динамические дашборды"]),
    ("Инструменты", ["Калькуляторы", "Справка и методики"]),
]

MENU_HANDLERS = {
    "Главная": show_home_page,
    "Активы": show_assets_page,
    "Угрозы": show_threats_page,
    "Уязвимости": show_vulnerabilities_page,
    "Контроли": show_controls_page,
    "Сценарии риска": show_risk_scenarios_page,
    "Реестр рисков": show_risk_register,
    "Риски": show_risks_page,
    "Мастер выбора модели": show_classification_wizard_page,
    "Оценка зрелости": show_maturity_page,
    "Экономика (EVI, ROI, TCO)": show_economic_page,
    "Аналитика": show_analytics_page,
    "Аналитика 3D": show_3d_visualization_page,
    "Динамические дашборды": show_dashboards_page,
    "Калькуляторы": show_calculators_page,
    "Справка и методики": show_reference_page,
}


def render_sidebar_nav():
    st.sidebar.markdown(
        f'<div class="rn-brand">🛡️ РискНавигатор<small>версия {APP_VERSION}</small></div>',
        unsafe_allow_html=True,
    )
    current = st.session_state.get("page", "Главная")
    for caption, items in NAV_GROUPS:
        st.sidebar.markdown(f'<div class="rn-navcap">{caption}</div>', unsafe_allow_html=True)
        for item in items:
            is_active = (item == current)
            if st.sidebar.button(
                item, key=f"nav_{item}",
                type="primary" if is_active else "secondary",
                width='stretch',
            ):
                st.session_state.page = item
                st.rerun()


def main():
    inject_css()
    st.session_state.setdefault("page", "Главная")
    render_sidebar_nav()
    render_header()

    handler = MENU_HANDLERS.get(st.session_state.page, show_home_page)
    handler()


if __name__ == "__main__":
    main()
