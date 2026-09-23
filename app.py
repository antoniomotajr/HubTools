"""
TECH TOOL HUB - FastAPI + JSON + Tailwind CSS
==============================================

Instalação:
    pip install fastapi uvicorn

Execução local:
    python app.py

Dados do usuário no Windows:
    %LOCALAPPDATA%\\TechToolHub

Acesse:
    http://127.0.0.1:8000

Swagger / documentação da API:
    http://127.0.0.1:8000/docs
"""

from __future__ import annotations

import fnmatch
import json
import re
import os
import shutil
import sqlite3
import socket
import sys
import subprocess
import tempfile
import threading
import webbrowser
from datetime import date
from pathlib import Path
from uuid import uuid4
from typing import Any
from urllib.parse import quote, urlparse

import uvicorn
from fastapi import FastAPI, HTTPException, Request, status
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel, Field


# ============================================================
# CONFIGURAÇÃO
# ============================================================

APP_TITLE = "TECH TOOL HUB"

# RESOURCE_DIR: arquivos somente leitura que acompanham o app/PyInstaller/MSIX.
# APP_DATA_DIR: arquivos que o usuário altera durante o uso.
SOURCE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", SOURCE_DIR))

if os.name == "nt":
    _local_app_data = os.environ.get("LOCALAPPDATA")
    LOCAL_APP_DATA_ROOT = (
        Path(_local_app_data)
        if _local_app_data
        else Path.home() / "AppData" / "Local"
    )
else:
    LOCAL_APP_DATA_ROOT = Path.home() / ".local" / "share"

APP_DATA_DIR = LOCAL_APP_DATA_ROOT / "TechToolHub"
APP_DATA_DIR.mkdir(parents=True, exist_ok=True)

# Mantido como alias para compatibilidade com trechos que tratam recursos.
BASE_DIR = RESOURCE_DIR

DATA_FILE = APP_DATA_DIR / "apps_data.json"
LINKS_FILE = APP_DATA_DIR / "links_data.json"
LOCAL_APPS_FILE = APP_DATA_DIR / "local_apps_data.json"
PROJECTS_FILE = APP_DATA_DIR / "projects_data.json"
USER_PROFILE_FILE = APP_DATA_DIR / "user_profile.json"
USER_MEDIA_DIR = APP_DATA_DIR / "user_media"
ICONS_DIR = APP_DATA_DIR / "icons"
CATALOG_MIGRATION_FILE = APP_DATA_DIR / ".apps_catalog_v1_done"
OFFICIAL_ICONS_MIGRATION_FILE = APP_DATA_DIR / ".official_icons_v1_done"
GIT_PROJECTS_FILE = APP_DATA_DIR / "git_projects.json"

# Workspace Explorer: raiz operacional padrão solicitada.
# Não é persistida como preferência do usuário.
WORKSPACE_DEFAULT_ROOT = Path(r"D:\\python\\CHATGPT")
WORKSPACE_EXCLUDED_FOLDERS = {
    ".git", ".idea", ".vscode", "__pycache__",
    "node_modules", "venv", ".venv",
}

# Caminhos selecionados no Git & GitHub ficam somente na sessão atual.
SESSION_GIT_PROJECTS_STATE: dict[str, Any] = {
    "selected_path": "",
    "projects": [],
}

# A marca é recurso do aplicativo e permanece somente leitura no MSIX.
BRAND_IMAGE_FILE = RESOURCE_DIR / "brand_logo.png"

DATA_LOCK = threading.RLock()
LINKS_LOCK = threading.RLock()

# Política de segurança v2.25.1:
# links adicionados pelo usuário existem somente durante a sessão atual.
SESSION_LINKS: list[dict[str, Any]] = []
LOCAL_APPS_LOCK = threading.RLock()
PROJECTS_LOCK = threading.RLock()
USER_PROFILE_LOCK = threading.RLock()
BUILD_LOCK = threading.RLock()
GIT_PROJECTS_LOCK = threading.RLock()

BUILD_STATE: dict[str, Any] = {
    "running": False,
    "status": "idle",
    "message": "Pronto para gerar o executável.",
    "started_at": "",
    "finished_at": "",
    "output_path": "",
    "log": [],
}

MAX_ICON_SIZE = 2 * 1024 * 1024  # 2 MB
MAX_PROFILE_IMAGE_SIZE = 5 * 1024 * 1024  # 5 MB
BRAND_CANDIDATE_PATHS = [
    BRAND_IMAGE_FILE,
    SOURCE_DIR / "brand_logo.png",
    SOURCE_DIR / "hub_de_ferramentas_tech_neon.png",
    SOURCE_DIR / "hub_logo.png",
]

README_CANDIDATE_PATHS = [
    RESOURCE_DIR / "README-MSIX.txt",
    SOURCE_DIR / "README-MSIX.txt",
    SOURCE_DIR / "README.txt",
]

BUILD_RELEASE_DIR = SOURCE_DIR / "release" / "exe"
BUILD_WORK_DIR = SOURCE_DIR / "build-dashboard"
ALLOWED_ICON_TYPES = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/webp": ".webp",
    "image/gif": ".gif",
}

app = FastAPI(
    title=APP_TITLE,
    description="Hub local para centralização e gerenciamento de ferramentas.",
    version="2.26.2",
)


@app.middleware("http")
async def security_no_store_middleware(request: Request, call_next):
    """Impede cache de respostas operacionais/sensíveis."""
    response = await call_next(request)

    sensitive_prefixes = (
        "/api/github/",
        "/api/git-dashboard/",
        "/api/links",
        "/api/user-profile",
        "/api/security-policy",
        "/api/workspace/explorer",
    )

    if request.url.path.startswith(sensitive_prefixes):
        response.headers["Cache-Control"] = "no-store, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"

    return response


# ============================================================
# CATÁLOGO DE APLICATIVOS LOCAIS
# ============================================================

LOCAL_APPS_CATALOG: list[dict[str, Any]] = [
    # Windows / Microsoft
    {
        "id": "word",
        "name": "Microsoft Word",
        "group": "Windows",
        "icon": "📝",
        "description": "Editor de documentos do Microsoft Office.",
        "exe_names": ["WINWORD.EXE"],
        "start_names": ["Microsoft Word", "Word"],
        "known_paths": [
            r"%ProgramFiles%\Microsoft Office\root\Office16\WINWORD.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\WINWORD.EXE",
            r"%ProgramFiles%\Microsoft Office\Office16\WINWORD.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\Office16\WINWORD.EXE",
        ],
    },
    {
        "id": "excel",
        "name": "Microsoft Excel",
        "group": "Windows",
        "icon": "📊",
        "description": "Planilhas, cálculos e análise de dados.",
        "exe_names": ["EXCEL.EXE"],
        "start_names": ["Microsoft Excel", "Excel"],
        "known_paths": [
            r"%ProgramFiles%\Microsoft Office\root\Office16\EXCEL.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\EXCEL.EXE",
            r"%ProgramFiles%\Microsoft Office\Office16\EXCEL.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\Office16\EXCEL.EXE",
        ],
    },
    {
        "id": "powerpoint",
        "name": "Microsoft PowerPoint",
        "group": "Windows",
        "icon": "📽️",
        "description": "Criação e apresentação de slides.",
        "exe_names": ["POWERPNT.EXE"],
        "start_names": ["Microsoft PowerPoint", "PowerPoint"],
        "known_paths": [
            r"%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE",
        ],
    },
    {
        "id": "outlook",
        "name": "Microsoft Outlook",
        "group": "Windows",
        "icon": "✉️",
        "description": "E-mail, calendário e contatos.",
        "exe_names": ["OUTLOOK.EXE"],
        "start_names": ["Microsoft Outlook", "Outlook"],
        "known_paths": [
            r"%ProgramFiles%\Microsoft Office\root\Office16\OUTLOOK.EXE",
            r"%ProgramFiles(x86)%\Microsoft Office\root\Office16\OUTLOOK.EXE",
        ],
    },
    {
        "id": "explorer",
        "name": "Explorador de Arquivos",
        "group": "Windows",
        "icon": "📁",
        "description": "Acesse arquivos, pastas e unidades do computador.",
        "exe_names": ["explorer.exe"],
        "start_names": ["File Explorer", "Explorador de Arquivos"],
    },
    {
        "id": "calculator",
        "name": "Calculadora",
        "group": "Windows",
        "icon": "🧮",
        "description": "Calculadora do Windows.",
        "exe_names": ["calc.exe"],
        "start_names": ["Calculator", "Calculadora"],
    },
    {
        "id": "wordpad",
        "name": "WordPad",
        "group": "Windows",
        "icon": "📄",
        "description": "Editor de texto formatado clássico do Windows.",
        "exe_names": ["write.exe"],
        "start_names": ["WordPad"],
    },
    {
        "id": "notepad",
        "name": "Bloco de Notas",
        "group": "Windows",
        "icon": "📃",
        "description": "Editor simples de arquivos de texto.",
        "exe_names": ["notepad.exe"],
        "start_names": ["Notepad", "Bloco de Notas"],
    },
    {
        "id": "paint",
        "name": "Paint",
        "group": "Windows",
        "icon": "🎨",
        "description": "Editor de imagens do Windows.",
        "exe_names": ["mspaint.exe"],
        "start_names": ["Paint"],
    },
    {
        "id": "snipping_tool",
        "name": "Ferramenta de Captura",
        "group": "Windows",
        "icon": "✂️",
        "description": "Captura de telas e recortes.",
        "exe_names": ["SnippingTool.exe", "snippingtool.exe"],
        "start_names": ["Snipping Tool", "Ferramenta de Captura", "Captura e Esboço"],
    },
    {
        "id": "terminal",
        "name": "Windows Terminal",
        "group": "Windows",
        "icon": "⌨️",
        "description": "Terminal moderno do Windows.",
        "exe_names": ["wt.exe"],
        "start_names": ["Terminal", "Windows Terminal"],
    },
    {
        "id": "powershell",
        "name": "Windows PowerShell",
        "group": "Windows",
        "icon": "⚡",
        "description": "Shell e automação do Windows.",
        "exe_names": ["powershell.exe"],
        "start_names": ["Windows PowerShell"],
    },
    {
        "id": "cmd",
        "name": "Prompt de Comando",
        "group": "Windows",
        "icon": "▣",
        "description": "Prompt de comandos clássico do Windows.",
        "exe_names": ["cmd.exe"],
        "start_names": ["Command Prompt", "Prompt de Comando"],
    },
    {
        "id": "task_manager",
        "name": "Gerenciador de Tarefas",
        "group": "Windows",
        "icon": "📈",
        "description": "Processos, desempenho e aplicativos em execução.",
        "exe_names": ["taskmgr.exe"],
        "start_names": ["Task Manager", "Gerenciador de Tarefas"],
    },
    {
        "id": "control_panel",
        "name": "Painel de Controle",
        "group": "Windows",
        "icon": "⚙️",
        "description": "Configurações clássicas do Windows.",
        "exe_names": ["control.exe"],
        "start_names": ["Control Panel", "Painel de Controle"],
    },
    {
        "id": "libreoffice",
        "name": "LibreOffice",
        "group": "Windows",
        "icon": "📚",
        "description": "Suíte de escritório LibreOffice: Writer, Calc, Impress e outros.",
        "exe_names": ["soffice.exe"],
        "start_names": ["LibreOffice"],
        "known_paths": [
            r"%ProgramFiles%\LibreOffice\program\soffice.exe",
            r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe",
        ],
    },

    # Apple para Windows
    {
        "id": "itunes",
        "name": "iTunes",
        "group": "Apple",
        "icon": "🎵",
        "description": "Biblioteca e gerenciamento de mídia da Apple.",
        "exe_names": ["iTunes.exe"],
        "start_names": ["iTunes"],
        "known_paths": [
            r"%ProgramFiles%\iTunes\iTunes.exe",
            r"%ProgramFiles(x86)%\iTunes\iTunes.exe",
        ],
    },
    {
        "id": "quicktime",
        "name": "QuickTime Player",
        "group": "Apple",
        "icon": "▶️",
        "description": "Reprodutor multimídia QuickTime.",
        "exe_names": ["QuickTimePlayer.exe"],
        "start_names": ["QuickTime Player", "QuickTime"],
        "known_paths": [
            r"%ProgramFiles%\QuickTime\QuickTimePlayer.exe",
            r"%ProgramFiles(x86)%\QuickTime\QuickTimePlayer.exe",
        ],
    },
    {
        "id": "icloud",
        "name": "iCloud",
        "group": "Apple",
        "icon": "☁️",
        "description": "Sincronização do iCloud no Windows.",
        "exe_names": ["iCloud.exe", "iCloudDrive.exe"],
        "start_names": ["iCloud"],
        "known_paths": [
            r"%ProgramFiles%\Common Files\Apple\Internet Services\iCloud.exe",
            r"%ProgramFiles(x86)%\Common Files\Apple\Internet Services\iCloud.exe",
            r"%ProgramFiles%\iCloud\iCloud.exe",
        ],
    },
    {
        "id": "apple_music",
        "name": "Apple Music",
        "group": "Apple",
        "icon": "🎧",
        "description": "Aplicativo Apple Music para Windows.",
        "exe_names": ["AppleMusic.exe"],
        "start_names": ["Apple Music"],
    },
    {
        "id": "apple_tv",
        "name": "Apple TV",
        "group": "Apple",
        "icon": "📺",
        "description": "Filmes, séries e conteúdo Apple TV.",
        "exe_names": ["AppleTV.exe"],
        "start_names": ["Apple TV"],
    },
    {
        "id": "apple_devices",
        "name": "Apple Devices",
        "group": "Apple",
        "icon": "📱",
        "description": "Gerenciamento de iPhone e iPad no Windows.",
        "exe_names": ["AppleDevices.exe"],
        "start_names": ["Apple Devices", "Dispositivos Apple"],
    },
]


WINDOWS_DIAGNOSTIC_BROWSERS: dict[str, dict[str, Any]] = {
    "chrome": {
        "name": "Google Chrome",
        "icon": "🌐",
        "exe_names": ["chrome.exe"],
        "known_paths": [
            r"%ProgramFiles%\Google\Chrome\Application\chrome.exe",
            r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe",
            r"%LocalAppData%\Google\Chrome\Application\chrome.exe",
        ],
        "cleanup_url": "chrome://settings/clearBrowserData",
    },
    "edge": {
        "name": "Microsoft Edge",
        "icon": "🌊",
        "exe_names": ["msedge.exe"],
        "known_paths": [
            r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe",
            r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe",
            r"%LocalAppData%\Microsoft\Edge\Application\msedge.exe",
        ],
        "cleanup_url": "edge://settings/clearBrowserData",
    },
    "firefox": {
        "name": "Mozilla Firefox",
        "icon": "🦊",
        "exe_names": ["firefox.exe"],
        "known_paths": [
            r"%ProgramFiles%\Mozilla Firefox\firefox.exe",
            r"%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe",
        ],
        "cleanup_url": "about:preferences#privacy",
    },
}



GIT_FORBIDDEN_PATTERNS = [
    ".env",
    ".env.*",
    "*.pfx",
    "*.p12",
    "*.pem",
    "*.key",
    "*.ppk",
    "apps_data.json",
    "links_data.json",
    "projects_data.json",
    "local_apps_data.json",
    "user_profile.json",
    "user_media/*",
    "icons/*",
    "release/*",
    "build/*",
    "dist/*",
    "build-msix/*",
    "build-dashboard/*",
    "*.exe",
    "*.msix",
    "*.msixbundle",
    "*.appx",
    "*.appxbundle",
    ".apps_catalog_v1_done",
    ".official_icons_v1_done",
]

GIT_WARNING_PATTERNS = [
    "*.cer",
    "*.log",
    "*.bak",
    "*.tmp",
    "*.spec",
]

SECRET_PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    (
        "private_key",
        re.compile(
            r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----",
            re.IGNORECASE,
        ),
    ),
    (
        "github_token",
        re.compile(
            r"\b(?:ghp_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,})\b"
        ),
    ),
    (
        "api_secret_assignment",
        re.compile(
            r"""(?ix)
            \b(
                api[_-]?key|
                client[_-]?secret|
                access[_-]?token|
                refresh[_-]?token|
                password|
                senha
            )\b
            \s*[:=]\s*
            ["'][^"']{8,}["']
            """
        ),
    ),
]

SECURITY_TEXT_EXTENSIONS = {
    ".py", ".ps1", ".js", ".ts", ".json", ".yml", ".yaml",
    ".toml", ".ini", ".cfg", ".conf", ".txt", ".md", ".html",
    ".css", ".xml", ".env",
}

SECURITY_SCAN_SKIP_NAMES = {
    "git-security-check.ps1",
    "GIT-SECURITY-AUDIT.md",
}

RELEASE_ALLOWED_SUFFIXES = {
    ".msix",
    ".msixbundle",
    ".appx",
    ".appxbundle",
    ".msi",
    ".exe",
    ".zip",
    ".7z",
    ".gz",
    ".whl",
    ".cer",
}

RELEASE_PRIVATE_SUFFIXES = {
    ".pfx",
    ".p12",
    ".pem",
    ".key",
    ".ppk",
}



def _find_github_cli() -> tuple[str, str]:
    """
    Localiza gh.exe sem depender apenas do PATH do processo.

    Isso é importante no MSIX/EXE, porque o aplicativo pode ter sido aberto
    antes da instalação do GitHub CLI ou pode receber um PATH diferente do
    PowerShell/Explorer.
    """
    found = shutil.which("gh")
    if found:
        return found, "PATH"

    if os.name != "nt":
        return "", ""

    candidate_specs: list[tuple[Path, str]] = []

    def add_candidate(raw: str | None, suffix: str, source: str) -> None:
        if not raw:
            return
        candidate_specs.append((Path(raw) / suffix, source))

    add_candidate(
        os.environ.get("ProgramFiles"),
        r"GitHub CLI\gh.exe",
        "Program Files",
    )
    add_candidate(
        os.environ.get("ProgramW6432"),
        r"GitHub CLI\gh.exe",
        "Program Files",
    )
    add_candidate(
        os.environ.get("ProgramFiles(x86)"),
        r"GitHub CLI\gh.exe",
        "Program Files (x86)",
    )
    add_candidate(
        os.environ.get("LOCALAPPDATA"),
        r"Programs\GitHub CLI\gh.exe",
        "LocalAppData Programs",
    )
    add_candidate(
        os.environ.get("LOCALAPPDATA"),
        r"Microsoft\WinGet\Links\gh.exe",
        "WinGet Links",
    )

    user_profile = os.environ.get("USERPROFILE")
    if user_profile:
        candidate_specs.append((
            Path(user_profile)
            / "AppData"
            / "Local"
            / "Microsoft"
            / "WinGet"
            / "Links"
            / "gh.exe",
            "WinGet Links",
        ))

    seen: set[str] = set()

    for candidate, source in candidate_specs:
        key = str(candidate).lower()
        if key in seen:
            continue
        seen.add(key)

        try:
            if candidate.is_file():
                return str(candidate), source
        except OSError:
            continue

    # Fallback para instalações do WinGet sem link atualizado no PATH.
    local_app_data = os.environ.get("LOCALAPPDATA")
    if local_app_data:
        packages_dir = (
            Path(local_app_data)
            / "Microsoft"
            / "WinGet"
            / "Packages"
        )
        try:
            if packages_dir.is_dir():
                patterns = (
                    "GitHub.cli_*",
                    "GitHub.cli*",
                )
                for pattern in patterns:
                    for package_dir in packages_dir.glob(pattern):
                        possible_files = [
                            package_dir / "gh.exe",
                            package_dir / "bin" / "gh.exe",
                        ]

                        try:
                            possible_files.extend(
                                list(package_dir.glob("**/gh.exe"))[:8]
                            )
                        except OSError:
                            pass

                        for candidate in possible_files:
                            try:
                                if candidate.is_file():
                                    return str(candidate), "WinGet Package"
                            except OSError:
                                continue
        except OSError:
            pass

    return "", ""


def _run_process(
    args: list[str],
    *,
    cwd: Path | None = None,
    timeout: int = 60,
) -> dict[str, Any]:
    """Executa um comando local sem shell e captura saída."""
    resolved_args = list(args)

    if resolved_args and resolved_args[0].lower() in {"gh", "gh.exe"}:
        gh_path, _gh_source = _find_github_cli()
        if gh_path:
            resolved_args[0] = gh_path

    creationflags = 0
    if os.name == "nt":
        creationflags = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))

    try:
        completed = subprocess.run(
            resolved_args,
            cwd=str(cwd or SOURCE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            shell=False,
            creationflags=creationflags,
        )
    except FileNotFoundError:
        return {
            "ok": False,
            "returncode": 127,
            "output": f"Comando não encontrado: {resolved_args[0]}",
        }
    except subprocess.TimeoutExpired as exc:
        output = exc.stdout or ""
        return {
            "ok": False,
            "returncode": 124,
            "output": f"{output}\nTempo limite excedido.".strip(),
        }

    return {
        "ok": completed.returncode == 0,
        "returncode": completed.returncode,
        "output": completed.stdout.strip(),
    }


def _clean_project_path(value: str | Path) -> Path | None:
    raw = str(value or "").strip().strip('"')
    if not raw:
        return None

    try:
        candidate = Path(os.path.expandvars(os.path.expanduser(raw))).resolve()
    except (OSError, RuntimeError):
        return None

    if not candidate.is_dir():
        return None

    return candidate


def _repo_root_from_path(project_path: str | Path) -> Path | None:
    """Resolve a raiz Git a partir de qualquer pasta interna do projeto."""
    project_dir = _clean_project_path(project_path)
    if project_dir is None or not shutil.which("git"):
        return None

    result = _run_process(
        ["git", "rev-parse", "--show-toplevel"],
        cwd=project_dir,
        timeout=15,
    )
    if not result["ok"] or not result["output"]:
        return None

    try:
        repo_root = Path(result["output"].splitlines()[-1]).resolve()
    except (OSError, RuntimeError):
        return None

    return repo_root if repo_root.is_dir() else None


def _default_git_projects_state() -> dict[str, Any]:
    source_root = _repo_root_from_path(SOURCE_DIR)
    if source_root is None:
        return {
            "selected_path": "",
            "projects": [],
        }

    return {
        "selected_path": str(source_root),
        "projects": [
            {
                "path": str(source_root),
                "name": source_root.name,
            }
        ],
    }


def _load_git_projects_state() -> dict[str, Any]:
    """Retorna projetos Git apenas da sessão atual."""
    with GIT_PROJECTS_LOCK:
        return {
            "selected_path": str(SESSION_GIT_PROJECTS_STATE.get("selected_path", "")).strip(),
            "projects": [
                {
                    "path": str(item.get("path", "")).strip(),
                    "name": str(item.get("name", "")).strip(),
                }
                for item in SESSION_GIT_PROJECTS_STATE.get("projects", [])
                if isinstance(item, dict) and str(item.get("path", "")).strip()
            ][:30],
        }



def _save_git_projects_state(state: dict[str, Any]) -> None:
    """Mantém o seletor Git apenas em memória; não grava caminhos locais."""
    global SESSION_GIT_PROJECTS_STATE
    with GIT_PROJECTS_LOCK:
        SESSION_GIT_PROJECTS_STATE = {
            "selected_path": str(state.get("selected_path", "")).strip(),
            "projects": [
                {
                    "path": str(item.get("path", "")).strip(),
                    "name": str(item.get("name", "")).strip(),
                }
                for item in state.get("projects", [])
                if isinstance(item, dict) and str(item.get("path", "")).strip()
            ][:30],
        }


def purge_legacy_git_project_paths() -> None:
    """Remove o arquivo legado que persistia caminhos locais de projetos Git."""
    try:
        GIT_PROJECTS_FILE.unlink(missing_ok=True)
    except OSError:
        pass



def _remember_git_project(
    project_path: str | Path,
    *,
    activate: bool = True,
) -> dict[str, Any]:
    candidate = _clean_project_path(project_path)
    if candidate is None:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A pasta selecionada não existe ou não pode ser acessada.",
        )

    repo_root = _repo_root_from_path(candidate)
    stored_path = repo_root or candidate
    state = _load_git_projects_state()

    project_item = {
        "path": str(stored_path),
        "name": stored_path.name,
    }

    projects = [
        item for item in state.get("projects", [])
        if str(item.get("path", "")).lower()
        != str(stored_path).lower()
    ]
    projects.insert(0, project_item)

    state["projects"] = projects[:30]
    if activate:
        state["selected_path"] = str(stored_path)

    _save_git_projects_state(state)

    return {
        "selected_path": str(stored_path),
        "repo_root": str(repo_root) if repo_root else "",
        "is_git": repo_root is not None,
        "name": stored_path.name,
    }


def _selected_project_path() -> Path | None:
    state = _load_git_projects_state()
    return _clean_project_path(
        str(state.get("selected_path", ""))
    )


def _git_repo_root() -> Path | None:
    selected = _selected_project_path()
    if selected is None:
        return None
    return _repo_root_from_path(selected)


def _is_source_git_mode() -> tuple[bool, str]:
    """Compatibilidade: agora valida o projeto selecionado, não só HubTools."""
    if not shutil.which("git"):
        return False, "Git não foi encontrado no PATH."

    selected = _selected_project_path()
    if selected is None:
        return False, "Nenhuma pasta de projeto está selecionada."

    repo_root = _repo_root_from_path(selected)
    if repo_root is None:
        return (
            False,
            "A pasta selecionada ainda não é um repositório Git.",
        )

    return True, ""


def _git_output(
    args: list[str],
    timeout: int = 30,
    *,
    repo_root: Path | None = None,
) -> str:
    root = repo_root or _git_repo_root()
    if root is None:
        return ""

    result = _run_process(
        ["git", *args],
        cwd=root,
        timeout=timeout,
    )
    return result["output"] if result["ok"] else ""


def _choose_project_folder() -> str:
    """Abre seletor nativo no Windows e retorna a pasta escolhida."""
    if os.name != "nt":
        raise HTTPException(
            status_code=status.HTTP_501_NOT_IMPLEMENTED,
            detail="O seletor nativo de pasta está disponível no Windows.",
        )

    powershell = (
        shutil.which("powershell.exe")
        or shutil.which("powershell")
    )
    if not powershell:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Windows PowerShell não foi encontrado.",
        )

    script = (
        "$ErrorActionPreference='Stop';"
        "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8;"
        "$shell=New-Object -ComObject Shell.Application;"
        "$folder=$shell.BrowseForFolder("
        "0,'Selecione a pasta do projeto',0,0);"
        "if($null -ne $folder){$folder.Self.Path}"
    )

    result = _run_process(
        [
            powershell,
            "-NoProfile",
            "-STA",
            "-Command",
            script,
        ],
        cwd=Path.home(),
        timeout=300,
    )

    if not result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=result["output"] or "Falha ao abrir o seletor de pasta.",
        )

    lines = [
        line.strip()
        for line in result["output"].splitlines()
        if line.strip()
    ]
    return lines[-1] if lines else ""


def _matches_any(path: str, patterns: list[str]) -> str | None:
    normalized = path.replace("\\", "/")
    basename = Path(normalized).name

    for pattern in patterns:
        normalized_pattern = pattern.replace("\\", "/")

        if "/" in normalized_pattern:
            if fnmatch.fnmatch(normalized, normalized_pattern):
                return pattern
        else:
            if (
                fnmatch.fnmatch(basename, normalized_pattern)
                or fnmatch.fnmatch(normalized, normalized_pattern)
            ):
                return pattern

    return None


def _git_candidate_files(repo_root: Path) -> list[str]:
    result = _run_process(
        [
            "git",
            "ls-files",
            "--cached",
            "--others",
            "--exclude-standard",
        ],
        cwd=repo_root,
        timeout=30,
    )
    if not result["ok"]:
        return []

    return sorted({
        line.strip().replace("\\", "/")
        for line in result["output"].splitlines()
        if line.strip()
    })


def _scan_file_for_secrets(
    repo_root: Path,
    relative_path: str,
) -> list[dict[str, Any]]:
    path = (repo_root / relative_path).resolve()

    try:
        path.relative_to(repo_root)
    except ValueError:
        return []

    if not path.is_file():
        return []

    if path.name in SECURITY_SCAN_SKIP_NAMES:
        return []

    if (
        path.suffix.lower() not in SECURITY_TEXT_EXTENSIONS
        and not path.name.startswith(".env")
    ):
        return []

    try:
        if path.stat().st_size > 2 * 1024 * 1024:
            return []
        content = path.read_text(
            encoding="utf-8",
            errors="replace",
        )
    except OSError:
        return []

    findings: list[dict[str, Any]] = []

    for line_number, line in enumerate(
        content.splitlines(),
        start=1,
    ):
        for finding_type, pattern in SECRET_PATTERNS:
            if not pattern.search(line):
                continue

            preview = line.strip()
            if len(preview) > 180:
                preview = preview[:177] + "..."

            findings.append({
                "severity": "high",
                "type": finding_type,
                "file": relative_path,
                "line": line_number,
                "message": (
                    "Possível segredo/credencial encontrado "
                    "no conteúdo."
                ),
                "preview": preview,
            })

    return findings


def run_git_security_audit(
    repo_root: Path | None = None,
) -> dict[str, Any]:
    """Audita candidatos do projeto atualmente selecionado."""
    available, reason = _is_source_git_mode()
    root = repo_root or _git_repo_root()

    if not available or root is None:
        return {
            "available": False,
            "reason": reason,
            "safe": False,
            "blocking_count": 0,
            "warning_count": 0,
            "findings": [],
            "candidate_count": 0,
            "tracked_count": 0,
        }

    candidates = _git_candidate_files(root)
    findings: list[dict[str, Any]] = []

    for relative_path in candidates:
        forbidden_pattern = _matches_any(
            relative_path,
            GIT_FORBIDDEN_PATTERNS,
        )

        if forbidden_pattern:
            findings.append({
                "severity": "high",
                "type": "forbidden_file",
                "file": relative_path,
                "line": None,
                "message": (
                    "Arquivo não deve ser publicado "
                    f"(regra: {forbidden_pattern})."
                ),
                "preview": "",
            })
            continue

        warning_pattern = _matches_any(
            relative_path,
            GIT_WARNING_PATTERNS,
        )

        if warning_pattern:
            findings.append({
                "severity": "warning",
                "type": "review_file",
                "file": relative_path,
                "line": None,
                "message": (
                    "Revisar antes de publicar "
                    f"(regra: {warning_pattern})."
                ),
                "preview": "",
            })

        findings.extend(
            _scan_file_for_secrets(
                root,
                relative_path,
            )
        )

    blocking_count = sum(
        1 for finding in findings
        if finding.get("severity") == "high"
    )
    warning_count = sum(
        1 for finding in findings
        if finding.get("severity") == "warning"
    )

    tracked_result = _run_process(
        ["git", "ls-files"],
        cwd=root,
        timeout=30,
    )
    tracked_count = len([
        line
        for line in tracked_result["output"].splitlines()
        if line.strip()
    ]) if tracked_result["ok"] else 0

    return {
        "available": True,
        "reason": "",
        "safe": blocking_count == 0,
        "blocking_count": blocking_count,
        "warning_count": warning_count,
        "candidate_count": len(candidates),
        "tracked_count": tracked_count,
        "findings": findings[:100],
    }


RELEASE_INTERNAL_NAMES = {
    "base_library.zip",
}

RELEASE_INTERNAL_DIR_NAMES = {
    "_internal",
    "__pycache__",
}


def _release_file_is_allowed(path: Path) -> bool:
    """Aceita apenas artefatos distribuíveis, nunca arquivos internos do build."""
    if not path.is_file():
        return False

    lower_name = path.name.lower()

    if lower_name in RELEASE_INTERNAL_NAMES:
        return False

    if any(
        part.lower() in RELEASE_INTERNAL_DIR_NAMES
        for part in path.parts
    ):
        return False

    if any(
        lower_name.endswith(suffix)
        for suffix in RELEASE_PRIVATE_SUFFIXES
    ):
        return False

    return any(
        lower_name.endswith(suffix)
        for suffix in RELEASE_ALLOWED_SUFFIXES
    )


def _release_asset_priority(
    item: dict[str, Any],
) -> tuple[int, float, str]:
    """
    Prioridade de seleção automática:
    0 = TechToolHub.exe / executável principal em release/exe
    1 = outros .exe
    2 = instaladores MSIX/MSI/AppX
    3 = pacotes ZIP/7z/whl
    9 = certificado
    """
    if item.get("is_certificate"):
        return (9, 0.0, str(item.get("name", "")).casefold())

    name = str(item.get("name", "")).casefold()
    relative = str(item.get("relative_path", "")).replace("\\", "/").casefold()
    suffix = Path(name).suffix.casefold()
    modified = float(item.get("modified", 0) or 0)

    if name == "techtoolhub.exe":
        return (0, -modified, name)

    if suffix == ".exe" and "/release/exe/" in f"/{relative}":
        return (0, -modified, name)

    if suffix == ".exe":
        return (1, -modified, name)

    if suffix in {
        ".msix",
        ".msixbundle",
        ".msi",
        ".appx",
        ".appxbundle",
    }:
        return (2, -modified, name)

    return (3, -modified, name)


def project_release_assets(
    repo_root: Path | None = None,
) -> dict[str, Any]:
    """Lista artefatos públicos disponíveis em <projeto>/release."""
    root = repo_root or _git_repo_root()

    if root is None:
        return {
            "directory": "",
            "files": [],
            "primary": "",
            "certificate": "",
            "suggested_tag": "",
            "suggested_title": "",
            "expected_executable": "",
        }

    release_dir = root / "release"
    files: list[dict[str, Any]] = []

    if release_dir.is_dir():
        for candidate in release_dir.rglob("*"):
            if not _release_file_is_allowed(candidate):
                continue

            try:
                relative = candidate.relative_to(root)
                size = candidate.stat().st_size
                modified = candidate.stat().st_mtime
            except OSError:
                continue

            files.append({
                "name": candidate.name,
                "path": str(candidate),
                "relative_path": str(relative).replace("\\", "/"),
                "size": size,
                "modified": modified,
                "is_certificate": candidate.suffix.lower() == ".cer",
            })

    files.sort(key=_release_asset_priority)

    primary_item = next(
        (
            item for item in files
            if not item.get("is_certificate")
        ),
        None,
    )
    certificate_item = next(
        (
            item for item in files
            if item.get("is_certificate")
        ),
        None,
    )

    suggested_tag = ""
    suggested_title = ""

    if primary_item:
        version_match = re.search(
            r"(?<!\d)(\d+\.\d+\.\d+)(?:\.\d+)?(?!\d)",
            str(primary_item["name"]),
        )
        if version_match:
            suggested_tag = f"v{version_match.group(1)}"
            suggested_title = (
                f"{root.name} {suggested_tag}"
            )

    return {
        "directory": str(release_dir),
        "files": files[:60],
        "primary": (
            str(primary_item["path"])
            if primary_item else ""
        ),
        "certificate": (
            str(certificate_item["path"])
            if certificate_item else ""
        ),
        "suggested_tag": suggested_tag,
        "suggested_title": suggested_title,
        "expected_executable": (
            str(root / "release" / "exe" / "TechToolHub" / "TechToolHub.exe")
            if (
                root / "release" / "exe" / "TechToolHub" / "TechToolHub.exe"
            ).is_file()
            else ""
        ),
    }


def _selected_projects_public_state() -> dict[str, Any]:
    state = _load_git_projects_state()
    selected_path = str(state.get("selected_path", ""))
    projects: list[dict[str, Any]] = []

    for item in state.get("projects", []):
        path = _clean_project_path(
            str(item.get("path", ""))
        )
        if path is None:
            continue

        repo_root = _repo_root_from_path(path)

        projects.append({
            "name": str(item.get("name") or path.name),
            "path": str(path),
            "is_git": repo_root is not None,
            "repo_root": str(repo_root) if repo_root else "",
            "selected": (
                str(path).lower() == selected_path.lower()
                if os.name == "nt"
                else str(path) == selected_path
            ),
        })

    return {
        "selected_path": selected_path,
        "projects": projects,
    }


def git_github_environment() -> dict[str, Any]:
    """Resumo do projeto selecionado, Git, origin e GitHub CLI."""
    source_available, source_reason = _is_source_git_mode()
    selected = _selected_project_path()
    repo_root = _git_repo_root()
    project_state = _selected_projects_public_state()

    branch = ""
    remote = ""
    status_lines: list[str] = []
    last_commit = ""
    upstream = ""
    ahead = 0
    behind = 0

    if repo_root:
        branch = _git_output(
            ["branch", "--show-current"],
            repo_root=repo_root,
        )
        remote = _git_output(
            ["remote", "get-url", "origin"],
            repo_root=repo_root,
        )
        status_lines = [
            line
            for line in _git_output(
                ["status", "--porcelain"],
                repo_root=repo_root,
            ).splitlines()
            if line.strip()
        ]
        last_commit = _git_output(
            ["log", "-1", "--format=%h|%s|%ci"],
            repo_root=repo_root,
        )
        upstream = _git_output(
            [
                "rev-parse",
                "--abbrev-ref",
                "--symbolic-full-name",
                "@{upstream}",
            ],
            repo_root=repo_root,
        )

        if branch:
            remote_ref = f"origin/{branch}"
            verify = _run_process(
                [
                    "git",
                    "show-ref",
                    "--verify",
                    "--quiet",
                    f"refs/remotes/{remote_ref}",
                ],
                cwd=repo_root,
                timeout=15,
            )
            if verify["ok"]:
                divergence = _run_process(
                    [
                        "git",
                        "rev-list",
                        "--left-right",
                        "--count",
                        f"HEAD...{remote_ref}",
                    ],
                    cwd=repo_root,
                    timeout=15,
                )
                if divergence["ok"]:
                    parts = divergence["output"].split()
                    if len(parts) >= 2:
                        try:
                            ahead = int(parts[0])
                            behind = int(parts[1])
                        except ValueError:
                            ahead = 0
                            behind = 0

    gh_path, gh_detection_source = _find_github_cli()
    gh_installed = bool(gh_path)
    gh_authenticated = False
    gh_auth_message = "GitHub CLI não encontrado."

    if gh_installed:
        auth_result = _run_process(
            ["gh", "auth", "status"],
            cwd=repo_root or selected or SOURCE_DIR,
            timeout=30,
        )
        gh_authenticated = auth_result["ok"]
        gh_auth_message = auth_result["output"]

    selected_name = ""
    if repo_root:
        selected_name = repo_root.name
    elif selected:
        selected_name = selected.name

    return {
        "source_mode": source_available,
        "source_reason": source_reason,
        "selected_project": {
            "name": selected_name,
            "path": str(selected) if selected else "",
            "repo_root": str(repo_root) if repo_root else "",
            "is_git": repo_root is not None,
        },
        "projects": project_state["projects"],
        "repo_root": str(repo_root) if repo_root else "",
        "branch": branch,
        "remote": redact_sensitive_text(remote),
        "changed_files": len(status_lines),
        "status_lines": status_lines[:30],
        "last_commit": last_commit,
        "upstream": upstream,
        "ahead": ahead,
        "behind": behind,
        "git_installed": bool(shutil.which("git")),
        "gh_installed": gh_installed,
        "gh_detection_source": gh_detection_source,
        "gh_authenticated": gh_authenticated,
        "gh_auth_message": gh_auth_message,
        "release_assets": project_release_assets(repo_root),
    }


def _remote_branch_exists(
    repo_root: Path,
    branch: str,
) -> bool:
    result = _run_process(
        [
            "git",
            "show-ref",
            "--verify",
            "--quiet",
            f"refs/remotes/origin/{branch}",
        ],
        cwd=repo_root,
        timeout=20,
    )
    return result["ok"]


def publish_project_to_github(
    commit_message: str,
) -> dict[str, Any]:
    """Audita, commita, sincroniza com origin e faz push."""
    environment = git_github_environment()

    if not environment["source_mode"]:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                environment["source_reason"]
                or "Git indisponível."
            ),
        )

    repo_root = Path(environment["repo_root"])
    branch = environment["branch"]
    remote = environment["remote"]

    if not branch:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "O projeto está em HEAD destacado. "
                "Selecione/crie uma branch antes de publicar."
            ),
        )

    if not remote:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "O remote 'origin' não está configurado "
                "neste projeto."
            ),
        )

    audit = run_git_security_audit(repo_root)

    if audit["blocking_count"] > 0:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                f"Publicação bloqueada: "
                f"{audit['blocking_count']} risco(s) "
                "de segurança encontrado(s)."
            ),
        )

    logs: list[str] = []

    add_result = _run_process(
        ["git", "add", "-A"],
        cwd=repo_root,
        timeout=60,
    )
    logs.append("$ git add -A")
    if add_result["output"]:
        logs.append(redact_sensitive_text(add_result["output"]))

    if not add_result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=(
                add_result["output"]
                or "git add falhou."
            ),
        )

    staged_audit = run_git_security_audit(repo_root)
    if staged_audit["blocking_count"] > 0:
        _run_process(
            ["git", "reset"],
            cwd=repo_root,
            timeout=30,
        )
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "O staging foi cancelado porque a "
                "auditoria encontrou "
                f"{staged_audit['blocking_count']} risco(s)."
            ),
        )

    diff_result = _run_process(
        ["git", "diff", "--cached", "--quiet"],
        cwd=repo_root,
        timeout=30,
    )

    committed = False
    if diff_result["returncode"] == 1:
        commit_result = _run_process(
            [
                "git",
                "commit",
                "-m",
                commit_message.strip(),
            ],
            cwd=repo_root,
            timeout=120,
        )
        logs.append(
            f"$ git commit -m {commit_message.strip()!r}"
        )
        if commit_result["output"]:
            logs.append(redact_sensitive_text(commit_result["output"]))

        if not commit_result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=(
                    commit_result["output"]
                    or "git commit falhou."
                ),
            )
        committed = True

    fetch_result = _run_process(
        ["git", "fetch", "origin"],
        cwd=repo_root,
        timeout=180,
    )
    logs.append("$ git fetch origin")
    if fetch_result["output"]:
        logs.append(fetch_result["output"])

    if not fetch_result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=(
                fetch_result["output"]
                or "git fetch origin falhou."
            ),
        )

    if _remote_branch_exists(repo_root, branch):
        divergence = _run_process(
            [
                "git",
                "rev-list",
                "--left-right",
                "--count",
                f"HEAD...origin/{branch}",
            ],
            cwd=repo_root,
            timeout=30,
        )

        remote_ahead = 0
        if divergence["ok"]:
            parts = divergence["output"].split()
            if len(parts) >= 2:
                try:
                    remote_ahead = int(parts[1])
                except ValueError:
                    remote_ahead = 0

        if remote_ahead > 0:
            rebase_result = _run_process(
                [
                    "git",
                    "pull",
                    "--rebase",
                    "origin",
                    branch,
                ],
                cwd=repo_root,
                timeout=180,
            )
            logs.append(
                f"$ git pull --rebase origin {branch}"
            )
            if rebase_result["output"]:
                logs.append(rebase_result["output"])

            if not rebase_result["ok"]:
                conflicts_result = _run_process(
                    [
                        "git",
                        "diff",
                        "--name-only",
                        "--diff-filter=U",
                    ],
                    cwd=repo_root,
                    timeout=30,
                )
                conflict_files = [
                    line.strip()
                    for line in conflicts_result["output"].splitlines()
                    if line.strip()
                ]

                _run_process(
                    ["git", "rebase", "--abort"],
                    cwd=repo_root,
                    timeout=30,
                )

                detail = (
                    "O remoto possui alterações e o rebase "
                    "gerou conflito. O rebase foi abortado "
                    "para preservar o projeto."
                )
                if conflict_files:
                    detail += (
                        " Arquivos em conflito: "
                        + ", ".join(conflict_files[:12])
                    )

                raise HTTPException(
                    status_code=status.HTTP_409_CONFLICT,
                    detail=detail,
                )

    push_result = _run_process(
        [
            "git",
            "push",
            "-u",
            "origin",
            branch,
        ],
        cwd=repo_root,
        timeout=180,
    )
    logs.append(
        f"$ git push -u origin {branch}"
    )
    if push_result["output"]:
        logs.append(redact_sensitive_text(push_result["output"]))

    if not push_result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=(
                push_result["output"]
                or "git push falhou."
            ),
        )

    return {
        "message": (
            f"Projeto {repo_root.name} enviado "
            "ao GitHub com sucesso."
        ),
        "project": repo_root.name,
        "committed": committed,
        "branch": branch,
        "remote": remote,
        "log": logs,
        "audit": staged_audit,
    }


def _validated_release_asset(
    repo_root: Path,
    asset_path: str,
) -> Path | None:
    raw = str(asset_path or "").strip()
    if not raw:
        return None

    try:
        candidate = Path(raw).resolve()
        release_root = (repo_root / "release").resolve()
        candidate.relative_to(release_root)
    except (OSError, RuntimeError, ValueError):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=(
                "O artefato da Release precisa estar "
                "dentro da pasta release do projeto."
            ),
        )

    if not _release_file_is_allowed(candidate):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O arquivo selecionado não é um artefato de Release permitido.",
        )

    if candidate.suffix.lower() == ".cer":
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Selecione um artefato principal; o .cer é anexado separadamente.",
        )

    return candidate




GITHUB_REPO_ACTIONS = {
    "clone",
    "select_local",
    "archive",
    "unarchive",
    "rename",
    "delete",
}


def _require_github_cli() -> None:
    gh_path, _gh_source = _find_github_cli()

    if not gh_path:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "GitHub CLI (gh) não foi localizado. "
                "Confirme no PowerShell com: gh --version. "
                "Se funcionar, reinicie o TECH TOOL HUB."
            ),
        )

    auth = _run_process(
        ["gh", "auth", "status"],
        cwd=_selected_project_path() or SOURCE_DIR,
        timeout=30,
    )
    if not auth["ok"]:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=(
                "GitHub CLI não está autenticado. "
                "Execute no PowerShell: gh auth login"
            ),
        )


def _validate_github_repo_slug(repo: str) -> str:
    value = repo.strip()
    if not re.fullmatch(r"[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+", value):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Identificador de repositório GitHub inválido.",
        )
    return value


def _validate_github_repo_name(name: str) -> str:
    value = name.strip()
    if not re.fullmatch(r"[A-Za-z0-9._-]+", value) or value in {".", ".."}:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=(
                "O nome do repositório pode conter apenas letras, números, "
                "ponto, hífen e sublinhado."
            ),
        )
    return value


def _github_current_login() -> str:
    _require_github_cli()
    result = _run_process(
        ["gh", "api", "user", "--jq", ".login"],
        cwd=_selected_project_path() or SOURCE_DIR,
        timeout=30,
    )
    if not result["ok"] or not result["output"].strip():
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=result["output"] or "Não foi possível identificar a conta GitHub.",
        )
    return result["output"].strip()


def _github_slug_from_remote(remote: str) -> str:
    value = remote.strip()
    patterns = [
        r"^https?://github\.com/([^/]+)/([^/?#]+?)(?:\.git)?/?$",
        r"^git@github\.com:([^/]+)/(.+?)(?:\.git)?$",
        r"^ssh://git@github\.com/([^/]+)/(.+?)(?:\.git)?/?$",
    ]

    for pattern in patterns:
        match = re.match(pattern, value, flags=re.IGNORECASE)
        if not match:
            continue

        owner = match.group(1).strip()
        name = match.group(2).strip()

        if name.lower().endswith(".git"):
            name = name[:-4]

        if owner and name:
            return f"{owner}/{name}".lower()

    return ""


def _github_local_repo_map() -> dict[str, dict[str, Any]]:
    """Mapeia owner/repo para clones locais conhecidos pelo Hub."""
    mapping: dict[str, dict[str, Any]] = {}
    state = _load_git_projects_state()
    selected = _selected_project_path()
    paths: list[Path] = []

    for item in state.get("projects", []):
        path = _clean_project_path(str(item.get("path", "")))
        if path is not None:
            paths.append(path)

    if selected is not None:
        paths.append(selected)

    seen: set[str] = set()

    for path in paths:
        repo_root = _repo_root_from_path(path)
        if repo_root is None:
            continue

        key = str(repo_root).lower() if os.name == "nt" else str(repo_root)
        if key in seen:
            continue
        seen.add(key)

        remote = _git_output(["remote", "get-url", "origin"], repo_root=repo_root)
        slug = _github_slug_from_remote(remote)
        if not slug:
            continue

        branch = _git_output(["branch", "--show-current"], repo_root=repo_root)
        changes = [
            line
            for line in _git_output(["status", "--porcelain"], repo_root=repo_root).splitlines()
            if line.strip()
        ]

        mapping[slug] = {
            "path": str(repo_root),
            "branch": branch.strip(),
            "changed_files": len(changes),
            "selected": (
                str(repo_root).lower() == str(selected).lower()
                if os.name == "nt" and selected
                else repo_root == selected
            ),
        }

    return mapping


def _github_api_json(
    endpoint: str,
    *,
    method: str = "GET",
    fields: list[tuple[str, str, bool]] | None = None,
    timeout: int = 90,
) -> Any:
    """Executa GitHub API via gh; typed=True usa -F."""
    _require_github_cli()

    args = ["gh", "api", "--method", method.upper(), endpoint]

    for key, value, typed in fields or []:
        args.extend(["-F" if typed else "-f", f"{key}={value}"])

    result = _run_process(
        args,
        cwd=_selected_project_path() or SOURCE_DIR,
        timeout=timeout,
    )

    if not result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=result["output"] or f"GitHub API falhou em {endpoint}.",
        )

    try:
        return json.loads(result["output"]) if result["output"] else {}
    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Resposta inválida do GitHub: {exc}",
        ) from exc


def list_github_repositories(scope: str = "owner") -> dict[str, Any]:
    """Lista repositórios e associa clones locais conhecidos."""
    _require_github_cli()

    normalized_scope = scope.strip().lower()
    if normalized_scope not in {"owner", "all"}:
        normalized_scope = "owner"

    login = _github_current_login()
    affiliation = (
        "owner"
        if normalized_scope == "owner"
        else "owner,collaborator,organization_member"
    )

    repositories: list[dict[str, Any]] = []

    for page in range(1, 6):
        endpoint = (
            "user/repos"
            f"?per_page=100&page={page}"
            "&sort=updated&direction=desc"
            f"&affiliation={affiliation}"
        )
        page_data = _github_api_json(endpoint)

        if not isinstance(page_data, list):
            break

        repositories.extend(page_data)

        if len(page_data) < 100:
            break

    local_map = _github_local_repo_map()
    items: list[dict[str, Any]] = []

    for repo in repositories:
        full_name = str(repo.get("full_name") or "")
        local = local_map.get(full_name.lower())

        permissions = repo.get("permissions")
        if not isinstance(permissions, dict):
            permissions = {}

        owner_data = repo.get("owner")
        owner_login = (
            str(owner_data.get("login") or "")
            if isinstance(owner_data, dict)
            else ""
        )

        visibility = str(repo.get("visibility") or "")
        if not visibility:
            visibility = "private" if repo.get("private") else "public"

        items.append({
            "id": repo.get("id"),
            "name": str(repo.get("name") or ""),
            "full_name": full_name,
            "owner": owner_login,
            "description": str(repo.get("description") or ""),
            "url": str(repo.get("html_url") or ""),
            "clone_url": str(repo.get("clone_url") or ""),
            "ssh_url": str(repo.get("ssh_url") or ""),
            "homepage": str(repo.get("homepage") or ""),
            "visibility": visibility,
            "private": bool(repo.get("private")),
            "archived": bool(repo.get("archived")),
            "fork": bool(repo.get("fork")),
            "template": bool(repo.get("is_template")),
            "default_branch": str(repo.get("default_branch") or ""),
            "language": str(repo.get("language") or ""),
            "stars": int(repo.get("stargazers_count") or 0),
            "forks": int(repo.get("forks_count") or 0),
            "open_issues": int(repo.get("open_issues_count") or 0),
            "size_kb": int(repo.get("size") or 0),
            "updated_at": str(repo.get("updated_at") or ""),
            "pushed_at": str(repo.get("pushed_at") or ""),
            "has_issues": bool(repo.get("has_issues", True)),
            "has_projects": bool(repo.get("has_projects", True)),
            "has_wiki": bool(repo.get("has_wiki", True)),
            "can_admin": bool(
                permissions.get("admin")
                or owner_login.lower() == login.lower()
            ),
            "can_push": bool(
                permissions.get("push")
                or permissions.get("maintain")
                or permissions.get("admin")
                or owner_login.lower() == login.lower()
            ),
            "local": local or None,
        })

    summary = {
        "total": len(items),
        "public": sum(1 for item in items if item["visibility"] == "public"),
        "private": sum(1 for item in items if item["visibility"] == "private"),
        "archived": sum(1 for item in items if item["archived"]),
        "forks": sum(1 for item in items if item["fork"]),
        "local": sum(1 for item in items if item["local"]),
    }

    return {
        "login": login,
        "scope": normalized_scope,
        "summary": summary,
        "repositories": items,
    }


def get_github_repository(repo: str) -> dict[str, Any]:
    slug = _validate_github_repo_slug(repo)
    data = _github_api_json(f"repos/{slug}")
    if not isinstance(data, dict):
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Resposta inválida ao consultar o repositório.",
        )
    return data


def create_github_repository(payload: GitHubRepoCreateRequest) -> dict[str, Any]:
    _require_github_cli()

    name = _validate_github_repo_name(payload.name)
    login = _github_current_login()

    repo_data = _github_api_json(
        "user/repos",
        method="POST",
        fields=[
            ("name", name, False),
            ("description", payload.description.strip(), False),
            ("private", "true" if payload.visibility == "private" else "false", True),
            ("auto_init", "true" if payload.add_readme else "false", True),
        ],
        timeout=120,
    )

    full_name = str(repo_data.get("full_name") or f"{login}/{name}")
    cloned_path = ""

    if payload.clone_after:
        parent = _clean_project_path(payload.parent_path)
        if parent is None:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail=(
                    "Repositório criado, mas a pasta do clone é inválida. "
                    f"Repositório: {full_name}"
                ),
            )

        destination = parent / name
        if destination.exists():
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=(
                    "Repositório criado no GitHub, mas a pasta local "
                    f"já existe: {destination}"
                ),
            )

        clone_result = _run_process(
            ["gh", "repo", "clone", full_name, name],
            cwd=parent,
            timeout=600,
        )
        if not clone_result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=(
                    f"Repositório {full_name} criado, mas o clone falhou:\n"
                    f"{clone_result['output']}"
                ),
            )

        repo_root = _repo_root_from_path(destination)
        if repo_root:
            _remember_git_project(repo_root, activate=True)
            cloned_path = str(repo_root)

    return {
        "message": f"Repositório {full_name} criado com sucesso.",
        "repo": full_name,
        "url": str(repo_data.get("html_url") or ""),
        "cloned_path": cloned_path,
    }


def edit_github_repository(payload: GitHubRepoEditRequest) -> dict[str, Any]:
    slug = _validate_github_repo_slug(payload.repo)
    current = get_github_repository(slug)

    current_visibility = str(
        current.get("visibility")
        or ("private" if current.get("private") else "public")
    )

    if (
        payload.visibility == "public"
        and current_visibility != "public"
        and not payload.confirm_public
    ):
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "Tornar um repositório público expõe o código. "
                "Confirmação explícita é necessária."
            ),
        )

    fields: list[tuple[str, str, bool]] = [
        ("description", payload.description.strip(), False),
        ("homepage", payload.homepage.strip(), False),
        ("visibility", payload.visibility, False),
        ("has_issues", "true" if payload.has_issues else "false", True),
        ("has_projects", "true" if payload.has_projects else "false", True),
        ("has_wiki", "true" if payload.has_wiki else "false", True),
    ]

    if payload.default_branch.strip():
        fields.append(("default_branch", payload.default_branch.strip(), False))

    updated = _github_api_json(
        f"repos/{slug}",
        method="PATCH",
        fields=fields,
        timeout=90,
    )

    return {
        "message": f"{slug} atualizado com sucesso.",
        "repo": str(updated.get("full_name") or slug),
        "url": str(updated.get("html_url") or ""),
    }


def _github_clone_repository(
    slug: str,
    parent_path: str,
    folder_name: str,
) -> dict[str, Any]:
    _require_github_cli()
    repo = _validate_github_repo_slug(slug)

    parent = _clean_project_path(parent_path)
    if parent is None:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Selecione uma pasta-pai válida para o clone.",
        )

    target_name = (
        _validate_clone_folder_name(folder_name)
        if folder_name.strip()
        else repo.split("/", 1)[1]
    )
    destination = parent / target_name

    if destination.exists():
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f"A pasta de destino já existe: {destination}",
        )

    result = _run_process(
        ["gh", "repo", "clone", repo, target_name],
        cwd=parent,
        timeout=600,
    )
    if not result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=result["output"] or "Clone falhou.",
        )

    repo_root = _repo_root_from_path(destination)
    if repo_root is None:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Clone concluído, mas o repositório local não foi identificado.",
        )

    remembered = _remember_git_project(repo_root, activate=True)

    return {
        "message": f"{repo} clonado e selecionado como projeto ativo.",
        "repo": repo,
        "path": str(repo_root),
        "selected_path": remembered.get("selected_path", ""),
    }


def manage_github_repository(payload: GitHubRepoActionRequest) -> dict[str, Any]:
    _require_github_cli()

    repo = _validate_github_repo_slug(payload.repo)
    action = payload.action.strip().lower()

    if action not in GITHUB_REPO_ACTIONS:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Ação de repositório GitHub não permitida.",
        )

    local_map = _github_local_repo_map()
    local = local_map.get(repo.lower())

    if action == "clone":
        return _github_clone_repository(
            repo,
            payload.parent_path,
            payload.folder_name,
        )

    if action == "select_local":
        if not local:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Nenhum clone local conhecido para este repositório.",
            )

        path = _clean_project_path(str(local.get("path") or ""))
        if path is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="A pasta local não está mais disponível.",
            )

        state = _remember_git_project(path, activate=True)
        return {
            "message": f"{repo} selecionado como projeto ativo.",
            "repo": repo,
            "path": str(path),
            "selected_path": state.get("selected_path", ""),
        }

    if action in {"archive", "unarchive"}:
        archived = action == "archive"
        updated = _github_api_json(
            f"repos/{repo}",
            method="PATCH",
            fields=[
                ("archived", "true" if archived else "false", True),
            ],
            timeout=90,
        )
        return {
            "message": f"{repo} {'arquivado' if archived else 'restaurado'}.",
            "repo": str(updated.get("full_name") or repo),
            "url": str(updated.get("html_url") or ""),
        }

    if action == "rename":
        new_name = _validate_github_repo_name(payload.new_name)

        if payload.confirmation.strip() != repo:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=(
                    "Para renomear, confirme digitando exatamente "
                    f"{repo}."
                ),
            )

        updated = _github_api_json(
            f"repos/{repo}",
            method="PATCH",
            fields=[("name", new_name, False)],
            timeout=90,
        )

        new_full_name = str(
            updated.get("full_name")
            or f"{repo.split('/', 1)[0]}/{new_name}"
        )
        local_remote_updated = False

        if local:
            local_path = _clean_project_path(str(local.get("path") or ""))
            repo_root = (
                _repo_root_from_path(local_path)
                if local_path is not None
                else None
            )
            if repo_root is not None:
                old_remote = _git_output(
                    ["remote", "get-url", "origin"],
                    repo_root=repo_root,
                )
                if (
                    old_remote
                    and _github_slug_from_remote(old_remote) == repo.lower()
                ):
                    if old_remote.startswith("git@"):
                        new_remote = f"git@github.com:{new_full_name}.git"
                    elif old_remote.startswith("ssh://"):
                        new_remote = f"ssh://git@github.com/{new_full_name}.git"
                    else:
                        new_remote = f"https://github.com/{new_full_name}.git"

                    set_remote = _run_process(
                        ["git", "remote", "set-url", "origin", new_remote],
                        cwd=repo_root,
                        timeout=30,
                    )
                    local_remote_updated = set_remote["ok"]

        return {
            "message": f"{repo} renomeado para {new_full_name}.",
            "repo": new_full_name,
            "url": str(updated.get("html_url") or ""),
            "local_remote_updated": local_remote_updated,
        }

    if action == "delete":
        if payload.confirmation.strip() != repo:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=(
                    "Exclusão permanente bloqueada. "
                    f"Digite exatamente {repo} para confirmar."
                ),
            )

        result = _run_process(
            ["gh", "repo", "delete", repo, "--yes"],
            cwd=_selected_project_path() or SOURCE_DIR,
            timeout=120,
        )
        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=(
                    result["output"]
                    or (
                        "Não foi possível excluir. A conta pode precisar "
                        "da permissão delete_repo no GitHub CLI."
                    )
                ),
            )

        return {
            "message": f"{repo} excluído permanentemente do GitHub.",
            "repo": repo,
        }

    raise HTTPException(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail="Ação não implementada.",
    )


GIT_INDIVIDUAL_ACTIONS = {
    "status",
    "init",
    "clone",
    "add",
    "commit",
    "switch",
    "branches",
    "branch_create",
    "branch_delete",
    "fetch",
    "pull_rebase",
    "push",
    "log",
    "diff",
    "remote",
    "rm_cached",
}


def _git_command_display(args: list[str]) -> str:
    """Forma legível do comando sem usar shell."""
    def quote_arg(value: str) -> str:
        if not value:
            return '""'
        if re.search(r"[\s\"']", value):
            return '"' + value.replace('"', '\\"') + '"'
        return value

    return " ".join(quote_arg(str(item)) for item in args)


def _require_selected_project() -> Path:
    selected = _selected_project_path()
    if selected is None:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Nenhuma pasta de projeto está selecionada.",
        )
    return selected


def _require_selected_repo() -> Path:
    selected = _require_selected_project()
    repo_root = _repo_root_from_path(selected)
    if repo_root is None:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "A pasta selecionada ainda não é um repositório Git. "
                "Use o comando Init ou selecione outro projeto."
            ),
        )
    return repo_root


def _validate_git_branch_name(
    repo_root: Path,
    branch: str,
) -> str:
    branch = branch.strip()

    if not branch:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Informe o nome da branch.",
        )

    result = _run_process(
        ["git", "check-ref-format", "--branch", branch],
        cwd=repo_root,
        timeout=15,
    )

    if not result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Nome de branch inválido: {branch}",
        )

    return branch


def _validate_clone_url(raw_url: str) -> str:
    clone_url = raw_url.strip()

    if not clone_url:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Informe a URL do repositório.",
        )

    # HTTPS sem credencial embutida.
    if clone_url.startswith("https://"):
        parsed = urlparse(clone_url)
        if (
            not parsed.hostname
            or parsed.username
            or parsed.password
            or parsed.query
            or parsed.fragment
        ):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail=(
                    "Use uma URL HTTPS limpa, sem usuário, senha, token, "
                    "query string ou fragmento."
                ),
            )
        return clone_url

    # SSH tradicional do GitHub/GitLab etc.
    if re.fullmatch(
        r"git@[A-Za-z0-9._-]+:[A-Za-z0-9._/-]+(?:\.git)?",
        clone_url,
    ):
        return clone_url

    if clone_url.startswith("ssh://"):
        parsed = urlparse(clone_url)
        if not parsed.hostname or parsed.password:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="URL SSH inválida.",
            )
        return clone_url

    raise HTTPException(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail=(
            "Use uma URL HTTPS ou SSH de Git válida. "
            "Tokens embutidos na URL não são aceitos."
        ),
    )


def _validate_clone_folder_name(
    raw_name: str,
) -> str:
    folder_name = raw_name.strip()

    if not folder_name:
        return ""

    if not re.fullmatch(r"[A-Za-z0-9._-]+", folder_name):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=(
                "O nome da pasta do clone pode conter apenas letras, "
                "números, ponto, hífen e sublinhado."
            ),
        )

    if folder_name in {".", ".."}:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome de pasta inválido.",
        )

    return folder_name


def _git_command_result(
    action: str,
    args: list[str],
    result: dict[str, Any],
    *,
    cwd: Path,
    message: str = "",
    extra: dict[str, Any] | None = None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {
        "ok": bool(result.get("ok")),
        "action": action,
        "command": redact_sensitive_text(_git_command_display(args)),
        "cwd": str(cwd),
        "returncode": int(result.get("returncode", 1)),
        "output": redact_sensitive_text(str(result.get("output", ""))),
        "message": redact_sensitive_text(message),
    }
    if extra:
        payload.update(extra)
    return payload


def execute_individual_git_command(
    payload: GitCommandRequest,
) -> dict[str, Any]:
    """Executa somente comandos Git previamente autorizados."""
    if not shutil.which("git"):
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail="Git não foi encontrado no PATH.",
        )

    action = payload.action.strip().lower()

    if action not in GIT_INDIVIDUAL_ACTIONS:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Comando Git não permitido pelo Dashboard.",
        )

    # --------------------------------------------------------
    # INIT: permitido em pasta selecionada ainda sem .git.
    # --------------------------------------------------------
    if action == "init":
        selected = _require_selected_project()
        existing_repo = _repo_root_from_path(selected)

        if existing_repo is not None:
            result = _run_process(
                ["git", "status", "--short", "--branch"],
                cwd=existing_repo,
                timeout=30,
            )
            return _git_command_result(
                action,
                ["git", "status", "--short", "--branch"],
                result,
                cwd=existing_repo,
                message="A pasta já é um repositório Git.",
                extra={"repo_root": str(existing_repo)},
            )

        args = ["git", "init", "-b", "main"]
        result = _run_process(args, cwd=selected, timeout=60)

        if not result["ok"]:
            # Compatibilidade com versões antigas do Git.
            args = ["git", "init"]
            result = _run_process(args, cwd=selected, timeout=60)
            if result["ok"]:
                branch_result = _run_process(
                    ["git", "branch", "-M", "main"],
                    cwd=selected,
                    timeout=30,
                )
                if branch_result["output"]:
                    result["output"] = (
                        result["output"]
                        + "\n"
                        + branch_result["output"]
                    ).strip()

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=result["output"] or "git init falhou.",
            )

        remembered = _remember_git_project(
            selected,
            activate=True,
        )
        repo_root = _repo_root_from_path(selected) or selected

        return _git_command_result(
            action,
            args,
            result,
            cwd=selected,
            message="Repositório Git inicializado.",
            extra={
                "repo_root": str(repo_root),
                "selected_project_changed": True,
                "selected_path": remembered.get("selected_path", ""),
            },
        )

    # --------------------------------------------------------
    # CLONE: destino escolhido separadamente.
    # --------------------------------------------------------
    if action == "clone":
        clone_url = _validate_clone_url(payload.value)
        folder_name = _validate_clone_folder_name(
            payload.folder_name
        )

        parent = _clean_project_path(payload.parent_path)
        if parent is None:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail=(
                    "Selecione uma pasta-pai existente para receber o clone."
                ),
            )

        if folder_name:
            destination = parent / folder_name
            if destination.exists():
                raise HTTPException(
                    status_code=status.HTTP_409_CONFLICT,
                    detail=(
                        f"A pasta de destino já existe: {destination}"
                    ),
                )
            args = [
                "git",
                "clone",
                clone_url,
                folder_name,
            ]
        else:
            destination = None
            args = [
                "git",
                "clone",
                clone_url,
            ]

        result = _run_process(
            args,
            cwd=parent,
            timeout=600,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=result["output"] or "git clone falhou.",
            )

        if destination is None:
            repo_name = clone_url.rstrip("/").split("/")[-1]
            if ":" in repo_name and clone_url.startswith("git@"):
                repo_name = repo_name.split(":")[-1]
            if repo_name.lower().endswith(".git"):
                repo_name = repo_name[:-4]
            destination = parent / repo_name

        repo_root = _repo_root_from_path(destination)
        if repo_root is None:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=(
                    "O clone terminou, mas o repositório de destino "
                    "não pôde ser identificado."
                ),
            )

        remembered = _remember_git_project(
            repo_root,
            activate=True,
        )

        return _git_command_result(
            action,
            args,
            result,
            cwd=parent,
            message="Repositório clonado e ativado.",
            extra={
                "repo_root": str(repo_root),
                "selected_project_changed": True,
                "selected_path": remembered.get("selected_path", ""),
            },
        )

    # A partir daqui é obrigatório um repositório existente.
    repo_root = _require_selected_repo()

    # --------------------------------------------------------
    # Comandos somente de leitura.
    # --------------------------------------------------------
    read_commands: dict[str, list[str]] = {
        "status": ["git", "status", "--short", "--branch"],
        "branches": ["git", "branch", "-vv", "--all"],
        "log": [
            "git",
            "log",
            "--oneline",
            "--decorate",
            "--graph",
            "-n",
            "30",
        ],
        "remote": ["git", "remote", "-v"],
    }

    if action in read_commands:
        args = read_commands[action]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=60,
        )
        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=result["output"] or f"{action} falhou.",
            )
        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
        )

    if action == "diff":
        worktree = _run_process(
            ["git", "diff", "--stat"],
            cwd=repo_root,
            timeout=60,
        )
        staged = _run_process(
            ["git", "diff", "--cached", "--stat"],
            cwd=repo_root,
            timeout=60,
        )

        output_parts = [
            "=== WORKTREE ===",
            worktree["output"] or "(sem alterações não staged)",
            "",
            "=== STAGED ===",
            staged["output"] or "(sem alterações staged)",
        ]

        result = {
            "ok": worktree["ok"] and staged["ok"],
            "returncode": 0 if worktree["ok"] and staged["ok"] else 1,
            "output": "\n".join(output_parts),
        }
        return _git_command_result(
            action,
            ["git", "diff", "--stat", "+", "git", "diff", "--cached", "--stat"],
            result,
            cwd=repo_root,
        )

    # --------------------------------------------------------
    # Auditoria em ações que podem publicar/registrar arquivos.
    # --------------------------------------------------------
    if action in {"add", "commit", "push"}:
        audit = run_git_security_audit(repo_root)
        if audit["blocking_count"] > 0:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=(
                    f"Comando bloqueado: a auditoria encontrou "
                    f"{audit['blocking_count']} risco(s) alto(s)."
                ),
            )

    if action == "add":
        pathspec = payload.value.strip()

        if pathspec:
            args = ["git", "add", "--", pathspec]
        else:
            args = ["git", "add", "-A"]

        result = _run_process(
            args,
            cwd=repo_root,
            timeout=120,
        )
        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=result["output"] or "git add falhou.",
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message="Arquivos adicionados ao staging.",
        )

    if action == "rm_cached":
        pathspec = payload.value.strip()

        if not pathspec:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Informe o arquivo ou pasta a remover do índice.",
            )

        args = [
            "git",
            "rm",
            "-r",
            "--cached",
            "--",
            pathspec,
        ]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=120,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=result["output"] or "git rm --cached falhou.",
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message=(
                "Item removido do índice Git. "
                "O arquivo local foi preservado."
            ),
        )

    if action == "commit":
        message = payload.message.strip()

        if not message:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="Informe a mensagem do commit.",
            )

        args = [
            "git",
            "commit",
            "-m",
            message,
        ]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=180,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=result["output"] or "git commit falhou.",
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message="Commit criado.",
        )

    if action in {"switch", "branch_create", "branch_delete"}:
        branch = _validate_git_branch_name(
            repo_root,
            payload.value,
        )

        if action == "switch":
            args = (
                ["git", "switch", "-c", branch]
                if payload.create
                else ["git", "switch", branch]
            )
        elif action == "branch_create":
            args = ["git", "branch", branch]
        else:
            current_branch = _git_output(
                ["branch", "--show-current"],
                repo_root=repo_root,
            ).strip()

            if branch == current_branch:
                raise HTTPException(
                    status_code=status.HTTP_409_CONFLICT,
                    detail=(
                        "Não é possível excluir a branch atualmente ativa. "
                        "Faça switch para outra branch primeiro."
                    ),
                )

            args = [
                "git",
                "branch",
                "-D" if payload.force else "-d",
                branch,
            ]

        result = _run_process(
            args,
            cwd=repo_root,
            timeout=120,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=result["output"] or f"git {action} falhou.",
            )

        messages = {
            "switch": "Branch alterada.",
            "branch_create": "Branch criada.",
            "branch_delete": "Branch excluída localmente.",
        }

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message=messages[action],
        )

    if action == "fetch":
        if not _git_output(
            ["remote", "get-url", "origin"],
            repo_root=repo_root,
        ):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="O remote 'origin' não está configurado.",
            )

        args = ["git", "fetch", "origin"]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=240,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=result["output"] or "git fetch origin falhou.",
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message="Referências remotas atualizadas.",
        )

    if action == "pull_rebase":
        branch = _git_output(
            ["branch", "--show-current"],
            repo_root=repo_root,
        ).strip()

        if not branch:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="HEAD destacado: selecione uma branch antes do pull.",
            )

        args = [
            "git",
            "pull",
            "--rebase",
            "origin",
            branch,
        ]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=300,
        )

        if not result["ok"]:
            conflicts = _run_process(
                [
                    "git",
                    "diff",
                    "--name-only",
                    "--diff-filter=U",
                ],
                cwd=repo_root,
                timeout=30,
            )
            conflict_files = [
                line.strip()
                for line in conflicts["output"].splitlines()
                if line.strip()
            ]

            # O painel usa modo conservador: não deixa o repo preso em rebase.
            _run_process(
                ["git", "rebase", "--abort"],
                cwd=repo_root,
                timeout=30,
            )

            detail = (
                result["output"]
                or "git pull --rebase falhou."
            )
            if conflict_files:
                detail += (
                    "\nRebase abortado automaticamente. "
                    "Conflitos: "
                    + ", ".join(conflict_files[:20])
                )

            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=detail,
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message="Pull --rebase concluído.",
        )

    if action == "push":
        branch = _git_output(
            ["branch", "--show-current"],
            repo_root=repo_root,
        ).strip()

        if not branch:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="HEAD destacado: selecione uma branch antes do push.",
            )

        if not _git_output(
            ["remote", "get-url", "origin"],
            repo_root=repo_root,
        ):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="O remote 'origin' não está configurado.",
            )

        args = [
            "git",
            "push",
            "-u",
            "origin",
            branch,
        ]
        result = _run_process(
            args,
            cwd=repo_root,
            timeout=300,
        )

        if not result["ok"]:
            raise HTTPException(
                status_code=status.HTTP_502_BAD_GATEWAY,
                detail=result["output"] or "git push falhou.",
            )

        return _git_command_result(
            action,
            args,
            result,
            cwd=repo_root,
            message="Push concluído.",
        )

    raise HTTPException(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail="Ação Git não implementada.",
    )


def create_github_release(
    payload: GitReleaseRequest,
) -> dict[str, Any]:
    """Cria Release do projeto selecionado usando gh autenticado."""
    environment = git_github_environment()

    if not environment["source_mode"]:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                environment["source_reason"]
                or "Git indisponível."
            ),
        )

    if not environment["gh_installed"]:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "GitHub CLI (gh) não está instalado. "
                "Instale com: winget install --id GitHub.cli"
            ),
        )

    if not environment["gh_authenticated"]:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail=(
                "GitHub CLI não está autenticado. "
                "Execute: gh auth login"
            ),
        )

    repo_root = Path(environment["repo_root"])
    audit = run_git_security_audit(repo_root)

    if audit["blocking_count"] > 0:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=(
                "Release bloqueada até a auditoria "
                "de segurança ficar sem riscos altos."
            ),
        )

    tag = payload.tag.strip()
    title = payload.title.strip()

    if not re.fullmatch(
        r"[A-Za-z0-9][A-Za-z0-9._+-]{0,79}",
        tag,
    ):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A tag contém caracteres inválidos.",
        )

    duplicate_version = re.fullmatch(
        r"(v?\d+\.\d+\.\d+(?:\.\d+)?)\1",
        tag,
        re.IGNORECASE,
    )
    if duplicate_version:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=(
                "A tag parece estar duplicada. "
                f"Use apenas {duplicate_version.group(1)}."
            ),
        )

    existing = _run_process(
        ["gh", "release", "view", tag],
        cwd=repo_root,
        timeout=30,
    )
    if existing["ok"]:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f"A release {tag} já existe.",
        )

    assets = project_release_assets(repo_root)
    selected_asset = _validated_release_asset(
        repo_root,
        payload.asset_path,
    )

    if selected_asset is None and assets["primary"]:
        selected_asset = Path(assets["primary"])

    command = [
        "gh",
        "release",
        "create",
        tag,
    ]

    if selected_asset is not None:
        command.append(str(selected_asset))

    certificate: Path | None = None
    if payload.include_certificate and assets["certificate"]:
        certificate = Path(assets["certificate"])
        if certificate.is_file():
            command.append(str(certificate))

    command.extend([
        "--title",
        title,
        "--notes",
        (
            payload.notes.strip()
            or f"Release {title}"
        ),
        "--target",
        environment["branch"] or "main",
    ])

    if payload.draft:
        command.append("--draft")

    result = _run_process(
        command,
        cwd=repo_root,
        timeout=180,
    )

    if not result["ok"]:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=(
                result["output"]
                or "Falha ao criar GitHub Release."
            ),
        )

    release_url = ""
    for line in result["output"].splitlines():
        if line.strip().startswith("https://"):
            release_url = line.strip()

    if not release_url:
        view_result = _run_process(
            [
                "gh",
                "release",
                "view",
                tag,
                "--json",
                "url",
                "--jq",
                ".url",
            ],
            cwd=repo_root,
            timeout=30,
        )
        if view_result["ok"]:
            release_url = view_result["output"].strip()

    return {
        "message": f"Release {tag} criada com sucesso.",
        "project": repo_root.name,
        "tag": tag,
        "title": title,
        "url": release_url,
        "asset": (
            str(selected_asset)
            if selected_asset is not None
            else ""
        ),
        "certificate": (
            str(certificate)
            if certificate is not None
            and certificate.is_file()
            else ""
        ),
        "draft": payload.draft,
        "output": redact_sensitive_text(result["output"]),
    }


def _local_request_only(request: Request) -> None:
    """Restringe ações de sistema ao computador local."""
    client_host = request.client.host if request.client else ""
    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Esta ação só pode ser executada localmente.",
        )


def read_project_readme() -> str:
    """Carrega o README distribuído com o Hub."""
    for candidate in README_CANDIDATE_PATHS:
        if not candidate.is_file():
            continue
        try:
            return candidate.read_text(encoding="utf-8-sig")
        except (OSError, UnicodeDecodeError):
            continue

    return """TECH TOOL HUB

Execução local:
    python app.py

Dados:
    %LOCALAPPDATA%\\TechToolHub

Criar Executável:
    Execute o Hub com Python na pasta do projeto e use o botão
    "Criar Executável" no Dashboard.

MSIX:
    Execute build-msix.ps1 na pasta do projeto.
"""


def build_environment_status() -> dict[str, Any]:
    """Indica se esta instância pode recompilar o projeto."""
    frozen = bool(getattr(sys, "frozen", False))
    reasons: list[str] = []

    if os.name != "nt":
        reasons.append("O build do executável está disponível no Windows.")

    if frozen:
        reasons.append(
            "O Hub está executando como aplicativo empacotado. "
            "Use 'python app.py' na pasta fonte para recompilar."
        )

    if not (SOURCE_DIR / "app.py").is_file():
        reasons.append(f"Arquivo fonte não encontrado: {SOURCE_DIR / 'app.py'}")

    if not (SOURCE_DIR / "brand_logo.png").is_file():
        reasons.append(f"Logo não encontrada: {SOURCE_DIR / 'brand_logo.png'}")

    return {
        "available": not reasons,
        "frozen": frozen,
        "source_dir": str(SOURCE_DIR),
        "output_dir": str(BUILD_RELEASE_DIR),
        "reasons": reasons,
    }


def _set_build_state(**changes: Any) -> None:
    with BUILD_LOCK:
        BUILD_STATE.update(changes)


def _append_build_log(line: str) -> None:
    clean = str(line or "").rstrip()
    if not clean:
        return
    with BUILD_LOCK:
        log = list(BUILD_STATE.get("log", []))
        log.append(clean)
        BUILD_STATE["log"] = log[-120:]


def public_build_state() -> dict[str, Any]:
    with BUILD_LOCK:
        data = {
            key: (list(value) if isinstance(value, list) else value)
            for key, value in BUILD_STATE.items()
        }
    data["environment"] = build_environment_status()
    return data


def _run_executable_build() -> None:
    """Executa PyInstaller em background e registra o log."""
    from datetime import datetime

    _set_build_state(
        running=True,
        status="running",
        message="Gerando TechToolHub.exe em modo enxuto...",
        started_at=datetime.now().isoformat(timespec="seconds"),
        finished_at="",
        output_path="",
        log=[],
    )

    try:
        environment = build_environment_status()
        if not environment["available"]:
            raise RuntimeError(" ".join(environment["reasons"]))

        required_pyinstaller = "6.22.3"

        version_probe = subprocess.run(
            [
                sys.executable,
                "-c",
                "import PyInstaller; print(PyInstaller.__version__)",
            ],
            cwd=str(SOURCE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )

        current_pyinstaller = (
            version_probe.stdout.strip()
            if version_probe.returncode == 0
            else ""
        )

        if current_pyinstaller != required_pyinstaller:
            _append_build_log(
                f"Atualizando PyInstaller: "
                f"{current_pyinstaller or 'não instalado'} -> "
                f"{required_pyinstaller}"
            )

            upgrade = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "--upgrade",
                    f"pyinstaller=={required_pyinstaller}",
                ],
                cwd=str(SOURCE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300,
            )

            for line in upgrade.stdout.splitlines():
                _append_build_log(line)

            if upgrade.returncode != 0:
                raise RuntimeError(
                    "Não foi possível atualizar o PyInstaller para "
                    f"{required_pyinstaller}."
                )

        jaraco_probe = subprocess.run(
            [
                sys.executable,
                "-c",
                (
                    "import jaraco.text, jaraco.functools, jaraco.context; "
                    "print('jaraco ok')"
                ),
            ],
            cwd=str(SOURCE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=30,
        )

        if jaraco_probe.returncode != 0:
            _append_build_log(
                "Instalando dependências jaraco necessárias ao pkg_resources..."
            )

            jaraco_install = subprocess.run(
                [
                    sys.executable,
                    "-m",
                    "pip",
                    "install",
                    "--upgrade",
                    "jaraco.text",
                    "jaraco.functools",
                    "jaraco.context",
                ],
                cwd=str(SOURCE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300,
            )

            for line in jaraco_install.stdout.splitlines():
                _append_build_log(line)

            if jaraco_install.returncode != 0:
                raise RuntimeError(
                    "Não foi possível instalar as dependências jaraco."
                )

        BUILD_RELEASE_DIR.mkdir(parents=True, exist_ok=True)
        BUILD_WORK_DIR.mkdir(parents=True, exist_ok=True)

        work_dir = BUILD_WORK_DIR / "work"
        spec_dir = BUILD_WORK_DIR / "spec"
        shutil.rmtree(work_dir, ignore_errors=True)
        shutil.rmtree(spec_dir, ignore_errors=True)
        work_dir.mkdir(parents=True, exist_ok=True)
        spec_dir.mkdir(parents=True, exist_ok=True)

        command = [
            sys.executable, "-m", "PyInstaller",
            "--noconfirm", "--clean", "--onedir", "--noconsole",
            "--name", "TechToolHub",
        ]

        icon_file = SOURCE_DIR / "packaging" / "TechToolHub.ico"
        if icon_file.is_file():
            command.extend(["--icon", str(icon_file)])

        command.extend([
            "--add-data", f"{SOURCE_DIR / 'brand_logo.png'};.",
        ])

        readme_file = SOURCE_DIR / "README-MSIX.txt"
        if readme_file.is_file():
            command.extend(["--add-data", f"{readme_file};."])

        command.extend([
            "--hidden-import", "uvicorn.logging",
            "--hidden-import", "uvicorn.loops.auto",
            "--hidden-import", "uvicorn.protocols.http.auto",
            "--hidden-import", "uvicorn.protocols.websockets.auto",
            "--hidden-import", "uvicorn.lifespan.on",

            # Compatibilidade setuptools/pkg_resources.
            "--hidden-import", "jaraco.text",
            "--hidden-import", "jaraco.functools",
            "--hidden-import", "jaraco.context",
            "--collect-submodules", "jaraco",
            "--collect-data", "jaraco.text",

            # O TECH TOOL HUB não usa estes pacotes. Em ambientes Python
            # grandes eles podem ser descobertos por hooks opcionais e
            # transformar o build em milhares de entradas desnecessárias.
            "--exclude-module", "IPython",
            "--exclude-module", "matplotlib",
            "--exclude-module", "numpy",
            "--exclude-module", "pandas",
            "--exclude-module", "scipy",
            "--exclude-module", "sklearn",
            "--exclude-module", "pytest",
            "--exclude-module", "sphinx",
            "--exclude-module", "docutils",
            "--exclude-module", "jedi",
            "--exclude-module", "parso",
            "--exclude-module", "nbformat",
            "--exclude-module", "notebook",
            "--exclude-module", "jupyter",
            "--exclude-module", "jupyter_core",
            "--exclude-module", "jupyter_client",
            "--exclude-module", "zmq",
            "--exclude-module", "pygame",
            "--exclude-module", "PyQt5",
            "--exclude-module", "PyQt6",
            "--exclude-module", "PySide2",
            "--exclude-module", "PySide6",
            "--exclude-module", "tkinter",
            "--exclude-module", "astroid",
            "--exclude-module", "pylint",
            "--exclude-module", "lxml",

            "--distpath", str(BUILD_RELEASE_DIR),
            "--workpath", str(work_dir),
            "--specpath", str(spec_dir),
            str(SOURCE_DIR / "app.py"),
        ])

        _append_build_log("Comando:")
        _append_build_log(" ".join(
            f'"{part}"' if " " in part else part
            for part in command
        ))

        process = subprocess.Popen(
            command,
            cwd=str(SOURCE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            bufsize=1,
        )

        assert process.stdout is not None
        for line in process.stdout:
            _append_build_log(line)

        return_code = process.wait()
        if return_code != 0:
            raise RuntimeError(f"PyInstaller terminou com código {return_code}.")

        exe_path = BUILD_RELEASE_DIR / "TechToolHub" / "TechToolHub.exe"
        if not exe_path.is_file():
            raise RuntimeError("TechToolHub.exe não foi localizado após o build.")

        _set_build_state(
            running=False,
            status="success",
            message="Executável criado com sucesso.",
            finished_at=datetime.now().isoformat(timespec="seconds"),
            output_path=str(exe_path),
        )
        _append_build_log(f"Concluído: {exe_path}")

    except Exception as exc:
        _append_build_log(f"ERRO: {exc}")
        _set_build_state(
            running=False,
            status="error",
            message=str(exc),
            finished_at=datetime.now().isoformat(timespec="seconds"),
        )


def migrate_legacy_user_data() -> None:
    """
    Migra, uma única vez e sem sobrescrever, dados da versão portátil
    (arquivos ao lado de app.py) para %LOCALAPPDATA%\\TechToolHub.
    """
    try:
        legacy_dir = SOURCE_DIR.resolve()
        target_dir = APP_DATA_DIR.resolve()
    except OSError:
        legacy_dir = SOURCE_DIR
        target_dir = APP_DATA_DIR

    if legacy_dir == target_dir:
        return

    legacy_files = [
        "apps_data.json",
        "local_apps_data.json",
        "projects_data.json",
        "user_profile.json",
        ".apps_catalog_v1_done",
        ".official_icons_v1_done",
    ]

    for filename in legacy_files:
        source = legacy_dir / filename
        destination = APP_DATA_DIR / filename
        if not source.is_file() or destination.exists():
            continue
        try:
            shutil.copy2(source, destination)
        except OSError:
            continue

    for dirname in ["icons", "user_media"]:
        source_dir = legacy_dir / dirname
        destination_dir = APP_DATA_DIR / dirname

        if not source_dir.is_dir():
            continue

        destination_dir.mkdir(parents=True, exist_ok=True)
        try:
            for item in source_dir.iterdir():
                destination = destination_dir / item.name
                if destination.exists():
                    continue
                if item.is_file():
                    shutil.copy2(item, destination)
        except OSError:
            continue


def write_custom_local_apps(items: list[dict[str, Any]]) -> None:
    """Grava os aplicativos locais adicionados pelo usuário."""
    temp_file = LOCAL_APPS_FILE.with_suffix(".tmp")

    try:
        with temp_file.open("w", encoding="utf-8") as file:
            json.dump(items, file, ensure_ascii=False, indent=2)

        os.replace(temp_file, LOCAL_APPS_FILE)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível gravar local_apps_data.json: {exc}",
        ) from exc


def write_user_profile(profile: dict[str, Any]) -> None:
    """Grava o perfil do usuário local."""
    temp_file = USER_PROFILE_FILE.with_suffix(".tmp")

    try:
        with temp_file.open("w", encoding="utf-8") as file:
            json.dump(profile, file, ensure_ascii=False, indent=2)

        os.replace(temp_file, USER_PROFILE_FILE)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível gravar user_profile.json: {exc}",
        ) from exc


def ensure_user_profile_file() -> None:
    """Cria o perfil local padrão."""
    if USER_PROFILE_FILE.exists():
        return

    write_user_profile({
        "name": "Usuário local",
        "role": "Administrador",
        "image": "",
        "image_filename": "",
    })


def read_user_profile() -> dict[str, Any]:
    """Lê o perfil do usuário local."""
    ensure_user_profile_file()

    try:
        with USER_PROFILE_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, dict):
            raise ValueError("user_profile.json precisa conter um objeto.")

        data.setdefault("name", "Usuário local")
        data.setdefault("role", "Administrador")
        data.setdefault("image", "")
        data.setdefault("image_filename", "")
        return data

    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"user_profile.json está inválido: {exc.msg}",
        ) from exc
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível ler user_profile.json: {exc}",
        ) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc


def public_user_profile(profile: dict[str, Any]) -> dict[str, str]:
    """Expõe somente os campos necessários ao frontend."""
    return {
        "name": str(profile.get("name", "Usuário local") or "Usuário local"),
        "role": str(profile.get("role", "Administrador") or "Administrador"),
        "image": str(profile.get("image", "") or ""),
    }


def ensure_brand_image_file() -> None:
    """Tenta localizar a arte da marca em caminhos conhecidos."""
    if BRAND_IMAGE_FILE.is_file():
        return

    for candidate in BRAND_CANDIDATE_PATHS:
        candidate_path = Path(candidate)
        if not candidate_path.is_file():
            continue
        try:
            if candidate_path.resolve() == BRAND_IMAGE_FILE.resolve():
                return
        except OSError:
            pass
        try:
            shutil.copyfile(candidate_path, BRAND_IMAGE_FILE)
            return
        except OSError:
            continue


def ensure_custom_local_apps_file() -> None:
    """Cria o banco de aplicativos locais personalizados."""
    if not LOCAL_APPS_FILE.exists():
        write_custom_local_apps([])


def read_custom_local_apps() -> list[dict[str, Any]]:
    """Lê os aplicativos locais personalizados."""
    ensure_custom_local_apps_file()

    try:
        with LOCAL_APPS_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, list):
            raise ValueError("local_apps_data.json precisa conter uma lista.")

        return [item for item in data if isinstance(item, dict)]

    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"local_apps_data.json está inválido: {exc.msg}",
        ) from exc
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível ler local_apps_data.json: {exc}",
        ) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc


def get_combined_local_apps_catalog() -> list[dict[str, Any]]:
    """Une o catálogo interno aos aplicativos cadastrados pelo usuário."""
    with LOCAL_APPS_LOCK:
        custom = read_custom_local_apps()

    return [*LOCAL_APPS_CATALOG, *custom]


def _registry_app_path(exe_name: str) -> str | None:
    """Procura executáveis registrados em 'App Paths' no Windows."""
    if os.name != "nt":
        return None

    try:
        import winreg
    except ImportError:
        return None

    subkey = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe_name}"
    roots = [winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE]
    views = [0]

    if hasattr(winreg, "KEY_WOW64_64KEY"):
        views.extend([winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY])

    for root in roots:
        for view in views:
            try:
                with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | view) as key:
                    value, _ = winreg.QueryValueEx(key, None)
                    if value and Path(value).exists():
                        return str(Path(value))
            except OSError:
                continue

    return None


def _start_apps_map() -> list[dict[str, str]]:
    """Lê atalhos do Menu Iniciar, inclusive apps Microsoft Store."""
    if os.name != "nt":
        return []

    command = (
        "Get-StartApps | "
        "Select-Object Name,AppID | "
        "ConvertTo-Json -Compress"
    )

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", command],
            capture_output=True,
            text=True,
            timeout=8,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return []

    if result.returncode != 0 or not result.stdout.strip():
        return []

    try:
        data = json.loads(result.stdout)
    except json.JSONDecodeError:
        return []

    if isinstance(data, dict):
        data = [data]

    output: list[dict[str, str]] = []
    for item in data if isinstance(data, list) else []:
        if not isinstance(item, dict):
            continue

        name = str(item.get("Name", "")).strip()
        app_id = str(item.get("AppID", "")).strip()

        if name and app_id:
            output.append({"name": name, "app_id": app_id})

    return output


def _match_start_app(item: dict[str, Any], start_apps: list[dict[str, str]]) -> str | None:
    expected = [str(name).casefold() for name in item.get("start_names", [])]

    for app_entry in start_apps:
        actual = app_entry["name"].casefold()

        if any(actual == name or name in actual for name in expected):
            return app_entry["app_id"]

    return None


def _resolve_local_app(
    item: dict[str, Any],
    start_apps: list[dict[str, str]] | None = None,
) -> dict[str, str] | None:
    """Resolve uma entrada do catálogo para um lançador local seguro."""
    if os.name != "nt":
        return None

    # Aplicativos cadastrados manualmente usam o caminho exato salvo.
    custom_path = str(item.get("custom_path", "")).strip()
    if custom_path:
        expanded_custom_path = Path(os.path.expandvars(custom_path)).expanduser()

        if expanded_custom_path.is_file():
            return {
                "kind": "exe",
                "target": str(expanded_custom_path),
                "source": "Adicionado por você",
            }

        return None

    # 1. Registro do Windows.
    for exe_name in item.get("exe_names", []):
        registered = _registry_app_path(exe_name)
        if registered:
            return {
                "kind": "exe",
                "target": registered,
                "source": "Aplicativo instalado",
            }

    # 2. PATH do processo.
    for exe_name in item.get("exe_names", []):
        found = shutil.which(exe_name)
        if found:
            return {
                "kind": "exe",
                "target": found,
                "source": "Sistema",
            }

    # 3. Caminhos conhecidos.
    for raw_path in item.get("known_paths", []):
        expanded = Path(os.path.expandvars(raw_path))
        if expanded.exists():
            return {
                "kind": "exe",
                "target": str(expanded),
                "source": "Aplicativo instalado",
            }

    # 4. Menu Iniciar / Microsoft Store.
    start_apps = start_apps if start_apps is not None else _start_apps_map()
    app_id = _match_start_app(item, start_apps)

    if app_id:
        return {
            "kind": "shell_app",
            "target": app_id,
            "source": "Menu Iniciar",
        }

    return None


def get_local_apps_status() -> list[dict[str, Any]]:
    """Retorna o catálogo enriquecido com o estado real de instalação."""
    start_apps = _start_apps_map()
    output: list[dict[str, Any]] = []

    for item in get_combined_local_apps_catalog():
        launcher = _resolve_local_app(item, start_apps)

        output.append({
            "id": item["id"],
            "name": item["name"],
            "group": item["group"],
            "icon": item["icon"],
            "description": item["description"],
            "installed": launcher is not None,
            "source": launcher["source"] if launcher else "Não encontrado",
            "custom": bool(item.get("custom", False)),
        })

    return output


def launch_catalog_app(app_id: str) -> dict[str, Any]:
    """Abre somente aplicativos previamente definidos no catálogo."""
    if os.name != "nt":
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="A abertura de aplicativos locais está disponível somente no Windows.",
        )

    item = next(
        (entry for entry in get_combined_local_apps_catalog() if entry["id"] == app_id),
        None,
    )

    if item is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Aplicativo local não reconhecido.",
        )

    launcher = _resolve_local_app(item)

    if launcher is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=f"{item['name']} não foi encontrado neste computador.",
        )

    try:
        if launcher["kind"] == "shell_app":
            subprocess.Popen(
                ["explorer.exe", rf"shell:AppsFolder\{launcher['target']}"],
                close_fds=True,
            )
        else:
            subprocess.Popen(
                [launcher["target"]],
                cwd=str(Path(launcher["target"]).parent),
                close_fds=True,
            )
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir {item['name']}: {exc}",
        ) from exc

    return {
        "message": f"{item['name']} aberto.",
        "id": app_id,
        "name": item["name"],
    }



# ============================================================
# DIAGNÓSTICO WINDOWS
# ============================================================

def require_local_windows(request: Request) -> None:
    """Restringe ações que controlam o PC ao próprio computador Windows."""
    client_host = request.client.host if request.client else ""

    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Esta ação só pode ser executada no próprio computador.",
        )

    if os.name != "nt":
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Esta ferramenta está disponível somente no Windows.",
        )


def resolve_windows_executable(
    exe_names: list[str],
    known_paths: list[str] | None = None,
) -> str | None:
    """Resolve um executável conhecido usando registro, PATH e caminhos padrão."""
    if os.name != "nt":
        return None

    for exe_name in exe_names:
        registered = _registry_app_path(exe_name)
        if registered:
            return registered

    for exe_name in exe_names:
        found = shutil.which(exe_name)
        if found:
            return found

    for raw_path in known_paths or []:
        expanded = Path(os.path.expandvars(raw_path))
        if expanded.is_file():
            return str(expanded)

    return None


def get_windows_diagnostics_status() -> dict[str, Any]:
    """Informa recursos disponíveis para a tela de diagnóstico."""
    browsers: list[dict[str, Any]] = []

    for browser_id, config in WINDOWS_DIAGNOSTIC_BROWSERS.items():
        executable = resolve_windows_executable(
            config["exe_names"],
            config.get("known_paths", []),
        )

        browsers.append({
            "id": browser_id,
            "name": config["name"],
            "icon": config["icon"],
            "installed": executable is not None,
        })

    return {
        "windows": os.name == "nt",
        "browsers": browsers,
        "tools": {
            "temp_locations": os.name == "nt",
            "performance_report": resolve_windows_executable(["perfmon.exe"]) is not None,
            "memory_diagnostic": resolve_windows_executable(["mdsched.exe"]) is not None,
        },
    }


DISK_CLEANUP_AREA_LABELS: dict[str, str] = {
    "windows_temp": "Windows Temp (Prefetch, %TEMP%, %TMP%)",
    "recycle_bin": "Lixeira",
    "windows_cache": "Caches/diagnósticos do Windows",
    "browser_cache": "Cache dos navegadores",
    "browser_history": "Histórico de navegação",
    "browser_sessions": "Cookies e sessões",
}


def _unique_existing_paths(paths: list[Path]) -> list[Path]:
    output: list[Path] = []
    seen: set[str] = set()

    for path in paths:
        try:
            expanded = Path(os.path.expandvars(str(path))).expanduser()
            normalized = os.path.normcase(os.path.abspath(str(expanded)))
        except (OSError, ValueError):
            continue

        if not normalized or normalized in seen:
            continue

        seen.add(normalized)

        if expanded.exists():
            output.append(expanded)

    return output


def _windows_temp_paths() -> list[Path]:
    windows_dir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    return _unique_existing_paths([
        windows_dir / "Prefetch",
        windows_dir / "Temp",
        Path(os.environ.get("TEMP", "")),
        Path(os.environ.get("TMP", "")),
    ])


def _windows_cache_paths() -> list[Path]:
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    program_data = Path(os.environ.get("PROGRAMDATA", r"C:\ProgramData"))

    return _unique_existing_paths([
        local / "CrashDumps",
        local / "Microsoft" / "Windows" / "WER" / "ReportArchive",
        local / "Microsoft" / "Windows" / "WER" / "ReportQueue",
        local / "Microsoft" / "Windows" / "INetCache",
        program_data / "Microsoft" / "Windows" / "WER" / "ReportArchive",
        program_data / "Microsoft" / "Windows" / "WER" / "ReportQueue",
    ])


def _chromium_profile_dirs(root: Path) -> list[Path]:
    if not root.is_dir():
        return []

    profiles: list[Path] = []

    default = root / "Default"
    if default.is_dir():
        profiles.append(default)

    try:
        for child in root.iterdir():
            if child.is_dir() and child.name.startswith("Profile "):
                profiles.append(child)
    except OSError:
        pass

    return profiles


def _browser_profile_roots() -> dict[str, list[Path]]:
    local = Path(os.environ.get("LOCALAPPDATA", ""))
    roaming = Path(os.environ.get("APPDATA", ""))

    chrome_root = local / "Google" / "Chrome" / "User Data"
    edge_root = local / "Microsoft" / "Edge" / "User Data"
    firefox_root = roaming / "Mozilla" / "Firefox" / "Profiles"

    chrome_profiles = _chromium_profile_dirs(chrome_root)
    edge_profiles = _chromium_profile_dirs(edge_root)

    firefox_profiles: list[Path] = []
    if firefox_root.is_dir():
        try:
            firefox_profiles = [p for p in firefox_root.iterdir() if p.is_dir()]
        except OSError:
            firefox_profiles = []

    return {
        "chrome": chrome_profiles,
        "edge": edge_profiles,
        "firefox": firefox_profiles,
    }


def _browser_cache_paths() -> list[Path]:
    profiles = _browser_profile_roots()
    paths: list[Path] = []

    for profile in profiles["chrome"] + profiles["edge"]:
        paths.extend([
            profile / "Cache",
            profile / "Code Cache",
            profile / "GPUCache",
            profile / "Service Worker" / "CacheStorage",
        ])

    for profile in profiles["firefox"]:
        paths.append(profile / "cache2")

    return _unique_existing_paths(paths)


def _browser_history_paths() -> list[Path]:
    profiles = _browser_profile_roots()
    paths: list[Path] = []

    for profile in profiles["chrome"] + profiles["edge"]:
        paths.extend([
            profile / "History",
            profile / "History-journal",
            profile / "Visited Links",
        ])

    for profile in profiles["firefox"]:
        paths.extend([
            profile / "places.sqlite",
            profile / "places.sqlite-shm",
            profile / "places.sqlite-wal",
        ])

    return _unique_existing_paths(paths)


def _browser_session_paths() -> list[Path]:
    profiles = _browser_profile_roots()
    paths: list[Path] = []

    for profile in profiles["chrome"] + profiles["edge"]:
        paths.extend([
            profile / "Cookies",
            profile / "Cookies-journal",
            profile / "Network" / "Cookies",
            profile / "Network" / "Cookies-journal",
            profile / "Sessions",
        ])

    for profile in profiles["firefox"]:
        paths.extend([
            profile / "cookies.sqlite",
            profile / "cookies.sqlite-shm",
            profile / "cookies.sqlite-wal",
            profile / "sessionstore.jsonlz4",
            profile / "sessionstore-backups",
        ])

    return _unique_existing_paths(paths)


def _path_size(path: Path) -> int:
    """Calcula tamanho sem seguir links simbólicos e ignorando acessos negados."""
    try:
        if path.is_symlink():
            return 0

        if path.is_file():
            return path.stat().st_size

        if not path.is_dir():
            return 0
    except OSError:
        return 0

    total = 0
    stack = [path]

    while stack:
        current = stack.pop()

        try:
            with os.scandir(current) as entries:
                for entry in entries:
                    try:
                        if entry.is_symlink():
                            continue

                        if entry.is_file(follow_symlinks=False):
                            total += entry.stat(follow_symlinks=False).st_size
                        elif entry.is_dir(follow_symlinks=False):
                            stack.append(Path(entry.path))
                    except OSError:
                        continue
        except OSError:
            continue

    return total


def _recycle_bin_size() -> int:
    if os.name != "nt":
        return 0

    script = r"""
$shell = New-Object -ComObject Shell.Application
$items = $shell.Namespace(10).Items()
$sum = ($items | Measure-Object -Property Size -Sum).Sum
if ($null -eq $sum) { $sum = 0 }
Write-Output ([int64]$sum)
"""

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            text=True,
            timeout=15,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        if result.returncode != 0:
            return 0

        return max(0, int(result.stdout.strip() or "0"))
    except (OSError, ValueError, subprocess.SubprocessError):
        return 0


def _disk_cleanup_area_paths(area_id: str) -> list[Path]:
    mapping = {
        "windows_temp": _windows_temp_paths,
        "windows_cache": _windows_cache_paths,
        "browser_cache": _browser_cache_paths,
        "browser_history": _browser_history_paths,
        "browser_sessions": _browser_session_paths,
    }

    resolver = mapping.get(area_id)
    return resolver() if resolver else []


def scan_disk_cleanup_areas(area_ids: list[str]) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    total = 0

    for area_id in area_ids:
        if area_id not in DISK_CLEANUP_AREA_LABELS:
            continue

        if area_id == "recycle_bin":
            size = _recycle_bin_size()
        else:
            size = sum(_path_size(path) for path in _disk_cleanup_area_paths(area_id))

        total += size
        results.append({
            "id": area_id,
            "label": DISK_CLEANUP_AREA_LABELS[area_id],
            "bytes": size,
        })

    return {
        "areas": results,
        "total_bytes": total,
    }


def _delete_known_path(path: Path) -> tuple[int, int]:
    """
    Limpa um arquivo ou o conteúdo de uma pasta conhecida.
    Retorna (bytes_removidos_estimados, itens_ignorados).
    """
    removed = 0
    skipped = 0

    try:
        if path.is_symlink():
            return (0, 1)

        if path.is_file():
            size = _path_size(path)
            try:
                path.unlink()
                return (size, 0)
            except OSError:
                return (0, 1)

        if not path.is_dir():
            return (0, 0)
    except OSError:
        return (0, 1)

    # Mantém a pasta raiz e remove somente seu conteúdo.
    try:
        children = list(path.iterdir())
    except OSError:
        return (0, 1)

    for child in children:
        try:
            child_size = _path_size(child)

            if child.is_symlink() or child.is_file():
                child.unlink()
                removed += child_size
                continue

            if child.is_dir():
                try:
                    shutil.rmtree(child)
                    removed += child_size
                except OSError:
                    # Fallback: remove o que estiver acessível dentro da subpasta.
                    sub_removed, sub_skipped = _delete_known_path(child)
                    removed += sub_removed
                    skipped += sub_skipped
                    try:
                        child.rmdir()
                    except OSError:
                        pass
        except OSError:
            skipped += 1

    return (removed, skipped)


def _clear_recycle_bin() -> tuple[int, int]:
    before = _recycle_bin_size()

    try:
        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-NonInteractive",
                "-Command",
                "Clear-RecycleBin -Force -ErrorAction Stop",
            ],
            capture_output=True,
            text=True,
            timeout=45,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )

        if result.returncode != 0:
            return (0, 1)

        return (before, 0)
    except (OSError, subprocess.SubprocessError):
        return (0, 1)


def clean_disk_cleanup_areas(area_ids: list[str]) -> dict[str, Any]:
    removed = 0
    skipped = 0
    cleaned_areas: list[str] = []

    for area_id in area_ids:
        if area_id not in DISK_CLEANUP_AREA_LABELS:
            continue

        area_removed = 0
        area_skipped = 0

        if area_id == "recycle_bin":
            area_removed, area_skipped = _clear_recycle_bin()
        else:
            for path in _disk_cleanup_area_paths(area_id):
                path_removed, path_skipped = _delete_known_path(path)
                area_removed += path_removed
                area_skipped += path_skipped

        removed += area_removed
        skipped += area_skipped
        cleaned_areas.append(area_id)

    return {
        "removed_bytes": removed,
        "skipped_items": skipped,
        "areas": cleaned_areas,
    }


def open_windows_temp_locations() -> list[str]:
    """Abre Prefetch, %TMP% e %TEMP% em janelas do Explorador."""
    if os.name != "nt":
        raise RuntimeError("Windows não disponível.")

    windows_dir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    candidates = [
        windows_dir / "Prefetch",
        Path(os.environ.get("TMP", "")),
        Path(os.environ.get("TEMP", "")),
    ]

    opened: list[str] = []
    seen: set[str] = set()

    for candidate in candidates:
        raw = str(candidate).strip()

        if not raw:
            continue

        normalized = os.path.normcase(os.path.abspath(raw))

        if normalized in seen:
            continue

        seen.add(normalized)

        if not Path(raw).exists():
            continue

        subprocess.Popen(
            ["explorer.exe", raw],
            close_fds=True,
        )
        opened.append(raw)

    return opened


def _shell_execute_windows(
    executable: str,
    args: tuple[str, ...] = (),
    *,
    run_as_admin: bool = False,
) -> None:
    """
    Inicia um executável usando ShellExecuteW.

    Quando run_as_admin=True, usa o verbo 'runas', que faz o Windows
    exibir o UAC para o próprio usuário confirmar a elevação.
    """
    if os.name != "nt":
        raise OSError("ShellExecuteW está disponível somente no Windows.")

    import ctypes

    parameters = subprocess.list2cmdline(list(args)) if args else None
    verb = "runas" if run_as_admin else "open"

    result = ctypes.windll.shell32.ShellExecuteW(
        None,
        verb,
        executable,
        parameters,
        None,
        1,  # SW_SHOWNORMAL
    )

    # ShellExecute retorna valor > 32 quando conseguiu iniciar.
    if int(result) <= 32:
        error_messages = {
            2: "arquivo não encontrado",
            3: "caminho não encontrado",
            5: "acesso negado ou elevação cancelada",
            8: "memória insuficiente",
            26: "falha de associação",
            27: "associação incompleta",
            28: "tempo excedido",
            31: "nenhum aplicativo associado",
        }
        message = error_messages.get(int(result), f"erro ShellExecute {int(result)}")
        raise OSError(f"Não foi possível iniciar o programa: {message}.")


def launch_windows_tool(
    executable_name: str,
    *args: str,
    elevated: bool = False,
) -> bool:
    """
    Abre uma ferramenta nativa conhecida do Windows.

    Retorna True quando a execução foi solicitada com elevação administrativa.
    Se uma execução normal gerar WinError 740, tenta novamente via UAC.
    """
    executable = resolve_windows_executable([executable_name])

    if not executable:
        raise FileNotFoundError(executable_name)

    if elevated:
        _shell_execute_windows(
            executable,
            tuple(args),
            run_as_admin=True,
        )
        return True

    try:
        subprocess.Popen(
            [executable, *args],
            close_fds=True,
        )
        return False

    except OSError as exc:
        # ERROR_ELEVATION_REQUIRED
        if getattr(exc, "winerror", None) == 740:
            _shell_execute_windows(
                executable,
                tuple(args),
                run_as_admin=True,
            )
            return True

        raise


def open_browser_cleanup(browser_id: str) -> str:
    """Abre a página nativa de limpeza do navegador escolhido."""
    config = WINDOWS_DIAGNOSTIC_BROWSERS.get(browser_id)

    if config is None:
        raise KeyError(browser_id)

    executable = resolve_windows_executable(
        config["exe_names"],
        config.get("known_paths", []),
    )

    if not executable:
        raise FileNotFoundError(config["name"])

    subprocess.Popen(
        [executable, config["cleanup_url"]],
        close_fds=True,
    )

    return config["name"]


# ============================================================
# MODELO DE DADOS
# ============================================================

class DiskCleanupRequest(BaseModel):
    areas: list[str] = Field(..., min_length=1)
    confirmed: bool = False


class CustomLocalAppCreate(BaseModel):
    name: str = Field(..., min_length=1, max_length=100)
    group: str = Field(default="Windows", pattern="^(Windows|Apple)$")
    executable_path: str = Field(..., min_length=3, max_length=1000)
    icon: str = Field(default="🖥️", min_length=1, max_length=16)
    description: str = Field(default="", max_length=240)


class AppItem(BaseModel):
    id: int = Field(..., gt=0, description="ID inteiro único")
    name: str = Field(..., min_length=1, max_length=80)
    category: str = Field(..., min_length=1, max_length=50)
    url: str = Field(..., min_length=1, max_length=500)
    icon: str = Field(..., min_length=1, max_length=500)
    description: str = Field(..., min_length=1, max_length=280)
    favorite: bool = Field(default=False, description="Indica se o aplicativo está favoritado")


class AppsOrderRequest(BaseModel):
    """Ordem persistente dos aplicativos exibidos na página APPS."""
    ids: list[int]


class FavoriteLinkCreate(BaseModel):
    title: str = Field(..., min_length=1, max_length=140)
    category: str = Field(..., min_length=1, max_length=60)
    url: str = Field(..., min_length=1, max_length=1000)
    note: str = Field(default="", max_length=300)


class FavoriteLinkItem(FavoriteLinkCreate):
    id: int = Field(..., gt=0)


class LinksOrderRequest(BaseModel):
    """Ordem persistente dos links favoritos."""
    ids: list[int]


class BrowserBookmarksImportRequest(BaseModel):
    browsers: list[str] = Field(..., min_length=1)
    use_folders: bool = True


class ProjectCreate(BaseModel):
    name: str = Field(..., min_length=1, max_length=120)
    client: str = Field(default="", max_length=120)
    start_date: str = Field(..., min_length=10, max_length=10)
    end_date: str = Field(default="", max_length=10)
    status: str = Field(
        default="em_andamento",
        pattern="^(planejado|em_andamento|pausado|concluido|cancelado)$",
    )
    notes: str = Field(default="", max_length=500)


class ProjectPhaseCreate(BaseModel):
    name: str = Field(..., min_length=1, max_length=100)
    status: str = Field(
        default="planejado",
        pattern="^(planejado|em_andamento|pausado|concluido)$",
    )
    progress: int = Field(default=0, ge=0, le=100)


class ProjectResourceRequest(BaseModel):
    resource_type: str = Field(..., pattern="^(app|local_app|link)$")
    resource_id: int | str


class ProjectResourcesOrderRequest(BaseModel):
    resources: list[ProjectResourceRequest]


class GitProjectPathRequest(BaseModel):
    path: str = Field(..., min_length=1, max_length=1200)


class WorkspaceFolderActionRequest(BaseModel):
    folder_name: str = Field(..., min_length=1, max_length=260)
    action: str = Field(..., pattern="^(open|terminal|git)$")


class GitPublishRequest(BaseModel):
    commit_message: str = Field(..., min_length=1, max_length=160)


class GitCommandRequest(BaseModel):
    action: str = Field(..., min_length=1, max_length=40)
    value: str = Field(default="", max_length=1600)
    message: str = Field(default="", max_length=240)
    create: bool = False
    force: bool = False
    parent_path: str = Field(default="", max_length=1600)
    folder_name: str = Field(default="", max_length=180)


class GitHubRepoCreateRequest(BaseModel):
    name: str = Field(..., min_length=1, max_length=100)
    description: str = Field(default="", max_length=350)
    visibility: str = Field(default="private", pattern="^(public|private)$")
    add_readme: bool = True
    clone_after: bool = False
    parent_path: str = Field(default="", max_length=1600)


class GitHubRepoEditRequest(BaseModel):
    repo: str = Field(..., min_length=3, max_length=220)
    description: str = Field(default="", max_length=350)
    homepage: str = Field(default="", max_length=500)
    visibility: str = Field(default="private", pattern="^(public|private|internal)$")
    has_issues: bool = True
    has_projects: bool = True
    has_wiki: bool = True
    default_branch: str = Field(default="", max_length=200)
    confirm_public: bool = False


class GitHubRepoActionRequest(BaseModel):
    repo: str = Field(..., min_length=3, max_length=220)
    action: str = Field(..., min_length=1, max_length=40)
    parent_path: str = Field(default="", max_length=1600)
    folder_name: str = Field(default="", max_length=180)
    new_name: str = Field(default="", max_length=100)
    confirmation: str = Field(default="", max_length=220)


class GitReleaseRequest(BaseModel):
    tag: str = Field(..., min_length=1, max_length=80)
    title: str = Field(..., min_length=1, max_length=140)
    notes: str = Field(default="", max_length=5000)
    asset_path: str = Field(default="", max_length=1600)
    include_certificate: bool = True
    draft: bool = False


# ============================================================
# DADOS INICIAIS
# ============================================================

DEFAULT_APPS = [
    {
        "id": 1,
        "name": "Notion",
        "category": "Produtividade",
        "url": "https://www.notion.so/",
        "icon": "📝",
        "description": "Organize anotações, projetos, documentos e bases de conhecimento.",
    },
    {
        "id": 2,
        "name": "GitHub",
        "category": "Dev",
        "url": "https://github.com/",
        "icon": "💻",
        "description": "Gerencie repositórios, código-fonte, issues e projetos de desenvolvimento.",
    },
    {
        "id": 3,
        "name": "Trello",
        "category": "Produtividade",
        "url": "https://trello.com/",
        "icon": "📋",
        "description": "Organize tarefas e fluxos de trabalho usando quadros e cartões.",
    },
    {
        "id": 4,
        "name": "ChatGPT",
        "category": "IA",
        "url": "https://chatgpt.com/",
        "icon": "🤖",
        "description": "Assistente de IA para pesquisa, escrita, programação e produtividade.",
    },
    {
        "id": 5,
        "name": "Figma",
        "category": "Design",
        "url": "https://www.figma.com/",
        "icon": "🎨",
        "description": "Criação colaborativa de interfaces, protótipos e sistemas de design.",
    },
    {
        "id": 6,
        "name": "Google Drive",
        "category": "Arquivos",
        "url": "https://drive.google.com/",
        "icon": "☁️",
        "description": "Armazene, organize e compartilhe documentos e arquivos na nuvem.",
    },
]


# Catálogo solicitado para esta versão. Os itens abaixo são adicionados
# uma única vez ao apps_data.json existente, sem apagar apps já cadastrados.
# Nomes repetidos na lista de origem (Threads, TikTok, Telegram e WhatsApp)
# são consolidados em um único cadastro.
REQUESTED_APPS = [
    {
        "name": "ChatGPT",
        "category": "IA",
        "url": "https://chatgpt.com/",
        "icon": "🤖",
        "description": "Assistente de IA para pesquisa, escrita, programação e produtividade.",
    },
    {
        "name": "Google Gemini",
        "category": "IA",
        "url": "https://gemini.google.com/",
        "icon": "✨",
        "description": "Assistente de IA do Google para pesquisa, criação e tarefas multimodais.",
    },
    {
        "name": "Threads",
        "category": "Redes Sociais",
        "url": "https://www.threads.net/",
        "icon": "🧵",
        "description": "Rede social de conversas e publicações conectada ao ecossistema da Meta.",
    },
    {
        "name": "CapCut",
        "category": "Vídeo e Criação",
        "url": "https://www.capcut.com/",
        "icon": "🎬",
        "description": "Editor de vídeo com recursos para criação, montagem e conteúdo para redes sociais.",
    },
    {
        "name": "Google Maps",
        "category": "Mapas",
        "url": "https://maps.google.com/",
        "icon": "🗺️",
        "description": "Mapas, rotas, localização de lugares e navegação.",
    },
    {
        "name": "Temu",
        "category": "Compras",
        "url": "https://www.temu.com/",
        "icon": "🛍️",
        "description": "Marketplace para pesquisa e compra de produtos online.",
    },
    {
        "name": "Google",
        "category": "Pesquisa",
        "url": "https://www.google.com/",
        "icon": "🔎",
        "description": "Mecanismo de busca para encontrar páginas, notícias, imagens e informações na web.",
    },
    {
        "name": "TikTok",
        "category": "Redes Sociais",
        "url": "https://www.tiktok.com/",
        "icon": "🎵",
        "description": "Plataforma de vídeos curtos, criação de conteúdo e descoberta de tendências.",
    },
    {
        "name": "Block Blast",
        "category": "Jogos",
        "url": "https://www.blockblast.com/",
        "icon": "🧩",
        "description": "Jogo de quebra-cabeça de blocos desenvolvido pela Hungry Studio.",
    },
    {
        "name": "Cici",
        "category": "IA",
        "url": "https://www.cici.com/",
        "icon": "💬",
        "description": "Assistente de IA para conversação, escrita, tradução e produtividade.",
    },
    {
        "name": "Google Chrome",
        "category": "Navegadores",
        "url": "https://www.google.com/chrome/",
        "icon": "🌐",
        "description": "Navegador web do Google para acesso a sites, serviços e aplicações online.",
    },
    {
        "name": "DeepSeek",
        "category": "IA",
        "url": "https://chat.deepseek.com/",
        "icon": "🧠",
        "description": "Assistente de IA para conversação, raciocínio, pesquisa e programação.",
    },
    {
        "name": "Telegram",
        "category": "Comunicação",
        "url": "https://web.telegram.org/",
        "icon": "✈️",
        "description": "Mensagens, grupos, canais e compartilhamento de arquivos.",
    },
    {
        "name": "Gmail",
        "category": "E-mail",
        "url": "https://mail.google.com/",
        "icon": "✉️",
        "description": "Serviço de e-mail do Google para mensagens, anexos e organização da caixa de entrada.",
    },
    {
        "name": "WhatsApp",
        "category": "Comunicação",
        "url": "https://web.whatsapp.com/",
        "icon": "🟢",
        "description": "Mensagens, chamadas e compartilhamento de arquivos pelo WhatsApp Web.",
    },
    {
        "name": "Instagram",
        "category": "Redes Sociais",
        "url": "https://www.instagram.com/",
        "icon": "📸",
        "description": "Rede social para fotos, vídeos, Stories, Reels e mensagens.",
    },
    {
        "name": "Facebook",
        "category": "Redes Sociais",
        "url": "https://www.facebook.com/",
        "icon": "👥",
        "description": "Rede social para publicações, comunidades, páginas e comunicação.",
    },
    {
        "name": "Snapchat",
        "category": "Redes Sociais",
        "url": "https://web.snapchat.com/",
        "icon": "👻",
        "description": "Aplicativo social de mensagens, fotos, vídeos e conteúdo temporário.",
    },
    {
        "name": "Pinterest",
        "category": "Redes Sociais",
        "url": "https://www.pinterest.com/",
        "icon": "📌",
        "description": "Plataforma visual para descobrir, organizar e salvar ideias e referências.",
    },
    {
        "name": "Messenger",
        "category": "Comunicação",
        "url": "https://www.messenger.com/",
        "icon": "💬",
        "description": "Serviço de mensagens da Meta para conversas e chamadas.",
    },
    {
        "name": "X",
        "category": "Redes Sociais",
        "url": "https://x.com/",
        "icon": "𝕏",
        "description": "Rede social para publicações, notícias, conversas e conteúdo em tempo real.",
    },
]


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def model_to_dict(model: BaseModel) -> dict[str, Any]:
    """Compatibilidade entre Pydantic v1 e v2."""
    if hasattr(model, "model_dump"):
        return model.model_dump()
    return model.dict()



SENSITIVE_OUTPUT_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (
        re.compile(r"\bghp_[A-Za-z0-9]{20,}\b"),
        "[REDACTED_GITHUB_TOKEN]",
    ),
    (
        re.compile(r"\bgithub_pat_[A-Za-z0-9_]{20,}\b"),
        "[REDACTED_GITHUB_TOKEN]",
    ),
    (
        re.compile(r"\b(?:sk|rk)-[A-Za-z0-9_-]{16,}\b"),
        "[REDACTED_TOKEN]",
    ),
    (
        re.compile(r"(?i)(Authorization\s*:\s*(?:Bearer|Basic)\s+)[^\s]+"),
        r"\1[REDACTED]",
    ),
    (
        re.compile(r"(?i)(https?://)([^/@\s:]+):([^/@\s]+)@"),
        r"\1[REDACTED]@",
    ),
]


def redact_sensitive_text(value: str) -> str:
    """Mascara tokens/credenciais antes de devolver texto ao frontend."""
    sanitized = str(value or "")
    for pattern, replacement in SENSITIVE_OUTPUT_PATTERNS:
        sanitized = pattern.sub(replacement, sanitized)
    return sanitized


def security_policy_payload() -> dict[str, Any]:
    """Política explícita de segurança do TECH TOOL HUB."""
    return {
        "zero_credentials": True,
        "stores_passwords": False,
        "stores_tokens": False,
        "stores_cookies": False,
        "stores_sessions": False,
        "stores_personal_links": False,
        "browser_authentication": "external",
        "message": (
            "O TECH TOOL HUB não armazena credenciais, tokens, cookies, "
            "sessões nem links pessoais. O acesso a sites e aplicativos "
            "depende da sessão existente no navegador ou no próprio aplicativo."
        ),
    }


def purge_legacy_personal_links_storage() -> None:
    """Remove o antigo links_data.json; a versão atual não usa esse arquivo."""
    try:
        LINKS_FILE.unlink(missing_ok=True)
    except OSError:
        # Mesmo se não for possível remover, a versão atual não lê o arquivo.
        pass


def validate_http_url(url: str) -> str:
    """Aceita somente URLs HTTP/HTTPS bem formadas."""
    clean_url = url.strip()
    parsed = urlparse(clean_url)

    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A URL deve ser válida e começar com http:// ou https://",
        )

    return clean_url


def official_icon_url(app_url: str) -> str:
    """
    Retorna uma URL de favicon em alta resolução para o site do aplicativo.

    O ícone é carregado a partir do domínio oficial informado no próprio
    cadastro do app. O navegador pode exibi-lo diretamente e, se o usuário
    enviar um ícone manualmente, o caminho local /icons/... tem prioridade.
    """
    clean_url = str(app_url or "").strip()
    if not clean_url:
        return ""

    return (
        "https://www.google.com/s2/favicons"
        f"?sz=128&domain_url={quote(clean_url, safe='')}"
    )


def migrate_official_icons(apps: list[dict[str, Any]]) -> bool:
    """
    Substitui emojis/ícones antigos pelos favicons oficiais dos sites.

    Regras:
    - não sobrescreve ícones enviados manualmente para /icons/;
    - não altera registros sem URL HTTP/HTTPS válida;
    - é executada como migração única para preservar alterações futuras.
    """
    changed = False

    for item in apps:
        if not isinstance(item, dict):
            continue

        current_icon = str(item.get("icon", "")).strip()
        app_url = str(item.get("url", "")).strip()

        # Ícone enviado pelo usuário: sempre preservar.
        if current_icon.startswith("/icons/"):
            continue

        parsed = urlparse(app_url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            continue

        new_icon = official_icon_url(app_url)
        if new_icon and current_icon != new_icon:
            item["icon"] = new_icon
            changed = True

    return changed


def write_apps(apps: list[dict[str, Any]]) -> None:
    """Grava o JSON de forma atômica para reduzir risco de corrupção."""
    temp_file = DATA_FILE.with_suffix(".tmp")

    try:
        with temp_file.open("w", encoding="utf-8") as file:
            json.dump(apps, file, ensure_ascii=False, indent=2)

        os.replace(temp_file, DATA_FILE)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível gravar apps_data.json: {exc}",
        ) from exc


def ensure_data_file() -> None:
    """Cria apps_data.json com dados de exemplo se ele ainda não existir."""
    if not DATA_FILE.exists():
        write_apps(DEFAULT_APPS)


def read_apps() -> list[dict[str, Any]]:
    """Lê todos os aplicativos do arquivo JSON."""
    ensure_data_file()

    try:
        with DATA_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, list):
            raise ValueError("O conteúdo raiz de apps_data.json precisa ser uma lista.")

        # Migração única do catálogo solicitado. Isso permite atualizar uma
        # instalação que já possua apps_data.json sem apagar os dados atuais.
        # Depois da primeira sincronização, o arquivo marcador impede que um
        # aplicativo excluído pelo usuário seja recriado em reinicializações futuras.
        if not CATALOG_MIGRATION_FILE.exists():
            existing_names = {
                str(item.get("name", "")).strip().casefold()
                for item in data
                if isinstance(item, dict)
            }
            next_id = max(
                (int(item.get("id", 0)) for item in data if isinstance(item, dict)),
                default=0,
            ) + 1
            changed = False

            for requested in REQUESTED_APPS:
                normalized_name = requested["name"].strip().casefold()
                if normalized_name in existing_names:
                    continue

                data.append({"id": next_id, **requested})
                existing_names.add(normalized_name)
                next_id += 1
                changed = True

            if changed:
                write_apps(data)

            try:
                CATALOG_MIGRATION_FILE.write_text(
                    "Catálogo solicitado sincronizado.\n",
                    encoding="utf-8",
                )
            except OSError:
                # Se o marcador não puder ser criado, a verificação poderá ocorrer
                # novamente, mas os nomes existentes impedem duplicação de apps.
                pass

        # Migração única de ícones oficiais. Ícones locais enviados manualmente
        # são preservados e nunca são substituídos por esta rotina.
        if not OFFICIAL_ICONS_MIGRATION_FILE.exists():
            icons_changed = migrate_official_icons(data)
            if icons_changed:
                write_apps(data)

            try:
                OFFICIAL_ICONS_MIGRATION_FILE.write_text(
                    "Ícones oficiais sincronizados.\n",
                    encoding="utf-8",
                )
            except OSError:
                # Se não for possível criar o marcador, a rotina poderá rodar
                # novamente; como a URL gerada é determinística, não duplica dados.
                pass

        return data

    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"O arquivo apps_data.json está inválido: {exc.msg}",
        ) from exc
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível ler apps_data.json: {exc}",
        ) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc


BROWSER_BOOKMARK_SOURCES: dict[str, dict[str, str]] = {
    "chrome": {
        "name": "Google Chrome",
        "kind": "chromium",
        "root": r"%LOCALAPPDATA%\Google\Chrome\User Data",
        "icon": "🌐",
    },
    "edge": {
        "name": "Microsoft Edge",
        "kind": "chromium",
        "root": r"%LOCALAPPDATA%\Microsoft\Edge\User Data",
        "icon": "🌊",
    },
    "brave": {
        "name": "Brave",
        "kind": "chromium",
        "root": r"%LOCALAPPDATA%\BraveSoftware\Brave-Browser\User Data",
        "icon": "🦁",
    },
    "firefox": {
        "name": "Mozilla Firefox",
        "kind": "firefox",
        "root": r"%APPDATA%\Mozilla\Firefox\Profiles",
        "icon": "🦊",
    },
}


def _safe_bookmark_category(value: str, fallback: str) -> str:
    category = " ".join(str(value or "").strip().split())
    if not category:
        category = fallback

    # FavoriteLinkCreate permite até 60 caracteres.
    return category[:60]


def _chromium_bookmark_profiles(browser_root: Path) -> list[Path]:
    """Localiza perfis Chromium com arquivo Bookmarks."""
    if not browser_root.is_dir():
        return []

    profiles: list[Path] = []

    for profile_name in ["Default", "Guest Profile"]:
        candidate = browser_root / profile_name
        if (candidate / "Bookmarks").is_file():
            profiles.append(candidate)

    try:
        for child in browser_root.iterdir():
            if (
                child.is_dir()
                and child.name.startswith("Profile ")
                and (child / "Bookmarks").is_file()
            ):
                profiles.append(child)
    except OSError:
        pass

    # Remove duplicações preservando a ordem.
    unique: list[Path] = []
    seen: set[str] = set()

    for profile in profiles:
        key = os.path.normcase(str(profile))
        if key not in seen:
            seen.add(key)
            unique.append(profile)

    return unique


def _walk_chromium_bookmark_node(
    node: dict[str, Any],
    *,
    browser_name: str,
    folder_path: list[str],
    output: list[dict[str, str]],
) -> None:
    node_type = str(node.get("type", ""))

    if node_type == "url":
        url = str(node.get("url", "")).strip()

        try:
            parsed = urlparse(url)
        except ValueError:
            return

        if parsed.scheme not in {"http", "https"}:
            return

        title = str(node.get("name", "")).strip()
        if not title:
            title = parsed.netloc or url

        category = folder_path[0] if folder_path else browser_name
        full_folder = " / ".join(folder_path)

        output.append({
            "title": title[:140],
            "url": url[:1000],
            "category": _safe_bookmark_category(category, browser_name),
            "note": (
                f"Importado do {browser_name}"
                + (f" • Pasta: {full_folder}" if full_folder else "")
            )[:300],
        })
        return

    if node_type == "folder":
        folder_name = str(node.get("name", "")).strip()
        next_path = [*folder_path]

        if folder_name:
            next_path.append(folder_name)

        for child in node.get("children", []) or []:
            if isinstance(child, dict):
                _walk_chromium_bookmark_node(
                    child,
                    browser_name=browser_name,
                    folder_path=next_path,
                    output=output,
                )


def read_chromium_bookmarks(
    browser_root: Path,
    browser_name: str,
) -> list[dict[str, str]]:
    """Lê favoritos dos perfis Chrome/Edge/Brave."""
    output: list[dict[str, str]] = []

    for profile in _chromium_bookmark_profiles(browser_root):
        bookmarks_file = profile / "Bookmarks"

        try:
            data = json.loads(bookmarks_file.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError, UnicodeDecodeError):
            continue

        roots = data.get("roots", {})
        if not isinstance(roots, dict):
            continue

        for root_key in ["bookmark_bar", "other", "synced"]:
            root = roots.get(root_key)

            if not isinstance(root, dict):
                continue

            for child in root.get("children", []) or []:
                if isinstance(child, dict):
                    _walk_chromium_bookmark_node(
                        child,
                        browser_name=browser_name,
                        folder_path=[],
                        output=output,
                    )

    return output


def _firefox_profiles(root: Path) -> list[Path]:
    if not root.is_dir():
        return []

    try:
        return [
            path for path in root.iterdir()
            if path.is_dir() and (path / "places.sqlite").is_file()
        ]
    except OSError:
        return []


def _read_firefox_places_db(db_path: Path) -> list[tuple[str, str, str]]:
    """
    Lê favoritos do Firefox. Tenta acesso somente leitura e,
    se necessário, consulta uma cópia temporária para evitar lock.
    """
    query = """
        SELECT
            COALESCE(NULLIF(TRIM(b.title), ''), NULLIF(TRIM(p.title), ''), p.url) AS title,
            p.url,
            COALESCE(parent.title, '') AS folder
        FROM moz_bookmarks AS b
        JOIN moz_places AS p ON p.id = b.fk
        LEFT JOIN moz_bookmarks AS parent ON parent.id = b.parent
        WHERE b.type = 1
          AND p.url IS NOT NULL
          AND (p.url LIKE 'http://%' OR p.url LIKE 'https://%')
        ORDER BY b.position ASC
    """

    def query_path(path: Path) -> list[tuple[str, str, str]]:
        connection = sqlite3.connect(
            f"file:{path.as_posix()}?mode=ro",
            uri=True,
            timeout=1.5,
        )
        try:
            rows = connection.execute(query).fetchall()
            return [
                (
                    str(row[0] or ""),
                    str(row[1] or ""),
                    str(row[2] or ""),
                )
                for row in rows
            ]
        finally:
            connection.close()

    try:
        return query_path(db_path)
    except (sqlite3.Error, OSError):
        pass

    temp_name = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".sqlite",
            delete=False,
        ) as temp_file:
            temp_name = temp_file.name

        shutil.copy2(db_path, temp_name)
        return query_path(Path(temp_name))
    except (sqlite3.Error, OSError):
        return []
    finally:
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                pass


def read_firefox_bookmarks(
    profiles_root: Path,
    browser_name: str,
) -> list[dict[str, str]]:
    """Lê favoritos do Firefox a partir de places.sqlite."""
    output: list[dict[str, str]] = []

    ignored_folders = {
        "menu",
        "toolbar",
        "unfiled",
        "mobile",
        "bookmarks menu",
        "bookmarks toolbar",
        "other bookmarks",
    }

    for profile in _firefox_profiles(profiles_root):
        for title, url, folder in _read_firefox_places_db(profile / "places.sqlite"):
            try:
                parsed = urlparse(url)
            except ValueError:
                continue

            if parsed.scheme not in {"http", "https"}:
                continue

            clean_folder = " ".join(folder.split())
            category = (
                clean_folder
                if clean_folder and clean_folder.casefold() not in ignored_folders
                else browser_name
            )

            output.append({
                "title": (title.strip() or parsed.netloc or url)[:140],
                "url": url[:1000],
                "category": _safe_bookmark_category(category, browser_name),
                "note": (
                    f"Importado do {browser_name}"
                    + (f" • Pasta: {clean_folder}" if clean_folder else "")
                )[:300],
            })

    return output


def read_browser_bookmarks(browser_id: str) -> list[dict[str, str]]:
    source = BROWSER_BOOKMARK_SOURCES.get(browser_id)
    if source is None:
        return []

    root = Path(os.path.expandvars(source["root"])).expanduser()
    browser_name = source["name"]

    if source["kind"] == "firefox":
        return read_firefox_bookmarks(root, browser_name)

    return read_chromium_bookmarks(root, browser_name)


def browser_bookmarks_status() -> list[dict[str, Any]]:
    """Conta favoritos encontrados nos navegadores suportados."""
    output: list[dict[str, Any]] = []

    for browser_id, source in BROWSER_BOOKMARK_SOURCES.items():
        bookmarks = read_browser_bookmarks(browser_id)

        output.append({
            "id": browser_id,
            "name": source["name"],
            "icon": source["icon"],
            "detected": bool(bookmarks),
            "count": len(bookmarks),
        })

    return output


def _workspace_explorer_root() -> Path:
    """Raiz do Workspace; não é salva como preferência."""
    if os.name == "nt":
        try:
            if WORKSPACE_DEFAULT_ROOT.is_dir():
                return WORKSPACE_DEFAULT_ROOT.resolve()
        except OSError:
            pass

    fallback = SOURCE_DIR.parent
    try:
        return fallback.resolve()
    except OSError:
        return fallback


def _workspace_direct_child(folder_name: str) -> Path:
    clean_name = str(folder_name or "").strip()

    if (
        not clean_name
        or clean_name in {".", ".."}
        or "/" in clean_name
        or "\\" in clean_name
    ):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome de pasta inválido.",
        )

    root = _workspace_explorer_root().resolve()
    candidate = (root / clean_name).resolve()

    try:
        candidate.relative_to(root)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Pasta fora da raiz do Workspace.",
        ) from exc

    if not candidate.is_dir() or candidate.parent != root:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Pasta do projeto não encontrada.",
        )

    return candidate


def _workspace_project_progress(project: dict[str, Any]) -> int:
    phases = [
        phase
        for phase in project.get("phases", [])
        if isinstance(phase, dict)
    ]

    if not phases:
        return 100 if project.get("status") == "concluido" else 0

    values: list[int] = []
    for phase in phases:
        try:
            value = int(phase.get("progress", 0))
        except (TypeError, ValueError):
            value = 0
        values.append(max(0, min(100, value)))

    return round(sum(values) / len(values)) if values else 0


def _workspace_folder_item_count(path: Path) -> int:
    try:
        return sum(1 for _ in path.iterdir())
    except OSError:
        return 0


def workspace_explorer_snapshot() -> dict[str, Any]:
    """Lista os diretórios reais da raiz como projetos estilo Explorer."""
    root = _workspace_explorer_root()

    with PROJECTS_LOCK:
        projects = read_projects()

    by_name = {
        str(project.get("name", "")).strip().casefold(): project
        for project in projects
        if isinstance(project, dict)
        and str(project.get("name", "")).strip()
    }

    try:
        children = sorted(
            (
                child
                for child in root.iterdir()
                if child.is_dir()
                and child.name not in WORKSPACE_EXCLUDED_FOLDERS
                and not child.name.startswith(".")
            ),
            key=lambda item: item.name.casefold(),
        )
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível ler a pasta do Workspace: {exc}",
        ) from exc

    items: list[dict[str, Any]] = []

    for folder in children:
        project = by_name.get(folder.name.casefold())
        try:
            modified_ts = float(folder.stat().st_mtime)
        except OSError:
            modified_ts = 0.0

        repo_root = _repo_root_from_path(folder)

        items.append({
            "name": folder.name,
            "modified_ts": modified_ts,
            "item_count": _workspace_folder_item_count(folder),
            "is_git": repo_root is not None,
            "registered": project is not None,
            "project_id": int(project.get("id")) if project and str(project.get("id", "")).isdigit() else None,
            "status": str(project.get("status") or "em_andamento") if project else "nao_classificado",
            "progress": _workspace_project_progress(project) if project else 0,
            "client": str(project.get("client") or "") if project else "",
        })

    summary = {
        "total": len(items),
        "em_andamento": sum(1 for item in items if item["status"] == "em_andamento"),
        "pausado": sum(1 for item in items if item["status"] == "pausado"),
        "concluido": sum(1 for item in items if item["status"] == "concluido"),
        "nao_classificado": sum(1 for item in items if item["status"] == "nao_classificado"),
        "git": sum(1 for item in items if item["is_git"]),
    }

    return {
        "root": str(root),
        "items": items,
        "summary": summary,
        "security": {
            "root_persisted": False,
            "git_paths_persisted": False,
        },
    }


def workspace_folder_action(payload: WorkspaceFolderActionRequest) -> dict[str, Any]:
    folder = _workspace_direct_child(payload.folder_name)

    if payload.action == "open":
        if os.name != "nt":
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Abrir no Explorer está disponível no Windows.",
            )

        try:
            subprocess.Popen(
                ["explorer.exe", str(folder)],
                cwd=str(folder),
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except OSError as exc:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Não foi possível abrir a pasta: {exc}",
            ) from exc

        return {
            "message": f"Pasta {folder.name} aberta no Explorer.",
            "folder_name": folder.name,
        }

    if payload.action == "terminal":
        if os.name != "nt":
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Terminal do Workspace está disponível no Windows.",
            )

        terminal = shutil.which("wt.exe")
        powershell = shutil.which("powershell.exe") or shutil.which("pwsh.exe")

        try:
            if terminal:
                subprocess.Popen(
                    [terminal, "-d", str(folder)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            elif powershell:
                subprocess.Popen(
                    [powershell, "-NoExit"],
                    cwd=str(folder),
                )
            else:
                raise OSError("Windows Terminal/PowerShell não encontrado.")
        except OSError as exc:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Não foi possível abrir o terminal: {exc}",
            ) from exc

        return {
            "message": f"Terminal aberto em {folder.name}.",
            "folder_name": folder.name,
        }

    if payload.action == "git":
        remembered = _remember_git_project(folder, activate=True)
        return {
            "message": (
                f"{folder.name} selecionado no Git & GitHub."
                if remembered.get("is_git")
                else f"{folder.name} selecionado. Use git init se quiser iniciar o repositório."
            ),
            "folder_name": folder.name,
            "is_git": bool(remembered.get("is_git")),
            "navigate_to": "/git-github",
        }

    raise HTTPException(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail="Ação do Workspace não suportada.",
    )


def default_project_phases() -> list[dict[str, Any]]:
    """Fases genéricas iniciais para um projeto novo ou legado."""
    names = ["Planejamento", "Execução", "Validação", "Entrega"]
    return [
        {
            "id": index,
            "name": name,
            "status": "planejado",
            "progress": 0,
        }
        for index, name in enumerate(names, start=1)
    ]


def write_projects(projects: list[dict[str, Any]]) -> None:
    """Grava projects_data.json de forma atômica."""
    temp_file = PROJECTS_FILE.with_suffix(".tmp")

    try:
        with temp_file.open("w", encoding="utf-8") as file:
            json.dump(projects, file, ensure_ascii=False, indent=2)

        os.replace(temp_file, PROJECTS_FILE)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível gravar projects_data.json: {exc}",
        ) from exc


def ensure_projects_file() -> None:
    """Cria o banco local de projetos."""
    if not PROJECTS_FILE.exists():
        write_projects([])


def read_projects() -> list[dict[str, Any]]:
    """Lê os projetos locais."""
    ensure_projects_file()

    try:
        with PROJECTS_FILE.open("r", encoding="utf-8") as file:
            data = json.load(file)

        if not isinstance(data, list):
            raise ValueError("projects_data.json precisa conter uma lista.")

        changed = False

        for project in data:
            if not isinstance(project, dict):
                continue

            if "resources" not in project:
                project["resources"] = []
                changed = True

            if "phases" not in project:
                project["phases"] = default_project_phases()
                changed = True

        cleaned = [project for project in data if isinstance(project, dict)]

        if changed:
            write_projects(cleaned)

        return cleaned

    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"projects_data.json está inválido: {exc.msg}",
        ) from exc
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível ler projects_data.json: {exc}",
        ) from exc
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(exc),
        ) from exc


def validate_project_dates(start_date: str, end_date: str = "") -> tuple[str, str]:
    """Valida datas ISO YYYY-MM-DD e a ordem início/fim."""
    try:
        start = date.fromisoformat(start_date)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Data de início inválida.",
        ) from exc

    clean_end = str(end_date or "").strip()
    if not clean_end:
        return start.isoformat(), ""

    try:
        end = date.fromisoformat(clean_end)
    except ValueError as exc:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Data final inválida.",
        ) from exc

    if end < start:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A data final não pode ser anterior à data de início.",
        )

    return start.isoformat(), end.isoformat()


def normalized_project_payload(payload: ProjectCreate) -> dict[str, Any]:
    """Normaliza dados textuais de um projeto."""
    start_date, end_date = validate_project_dates(
        payload.start_date.strip(),
        payload.end_date.strip(),
    )

    name = payload.name.strip()
    client = payload.client.strip()
    notes = payload.notes.strip()

    if not name:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O nome do projeto é obrigatório.",
        )

    return {
        "name": name,
        "client": client,
        "start_date": start_date,
        "end_date": end_date,
        "status": payload.status,
        "notes": notes,
    }


def normalize_project_resource_id(
    resource_type: str,
    resource_id: int | str,
) -> int | str:
    """Normaliza IDs numéricos de Apps/Links e IDs textuais de Apps Locais."""
    if resource_type in {"app", "link"}:
        try:
            normalized = int(resource_id)
        except (TypeError, ValueError) as exc:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="ID de recurso inválido.",
            ) from exc

        if normalized <= 0:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="ID de recurso inválido.",
            )

        return normalized

    if resource_type == "local_app":
        normalized = str(resource_id).strip()
        if not normalized:
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="ID do aplicativo local é obrigatório.",
            )
        return normalized

    raise HTTPException(
        status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
        detail="Tipo de recurso inválido.",
    )


def project_resource_key(
    resource_type: str,
    resource_id: int | str,
) -> tuple[str, str]:
    """Chave estável para comparar recursos de projetos."""
    normalized = normalize_project_resource_id(resource_type, resource_id)
    return resource_type, str(normalized)


def project_resource_exists(
    resource_type: str,
    resource_id: int | str,
) -> bool:
    """Confere se App, App Local ou Link ainda existe em seu banco de origem."""
    normalized = normalize_project_resource_id(resource_type, resource_id)

    if resource_type == "app":
        with DATA_LOCK:
            return any(
                int(item.get("id", -1)) == normalized
                for item in read_apps()
            )

    if resource_type == "link":
        with LINKS_LOCK:
            return any(
                int(item.get("id", -1)) == normalized
                for item in read_links()
            )

    if resource_type == "local_app":
        return any(
            str(item.get("id", "")) == str(normalized)
            for item in get_combined_local_apps_catalog()
        )

    return False


def write_links(links: list[dict[str, Any]]) -> None:
    """Mantém os links somente na memória da sessão atual."""
    global SESSION_LINKS
    SESSION_LINKS = [
        dict(item)
        for item in links
        if isinstance(item, dict)
    ]


def ensure_links_file() -> None:
    """Compatibilidade: nenhum arquivo de links é criado."""
    return


def read_links() -> list[dict[str, Any]]:
    """Retorna uma cópia dos links temporários da sessão."""
    return [
        dict(item)
        for item in SESSION_LINKS
        if isinstance(item, dict)
    ]


migrate_legacy_user_data()
purge_legacy_personal_links_storage()
purge_legacy_git_project_paths()

ICONS_DIR.mkdir(parents=True, exist_ok=True)

with DATA_LOCK:
    ensure_data_file()

with LINKS_LOCK:
    ensure_links_file()

with PROJECTS_LOCK:
    ensure_projects_file()

with LOCAL_APPS_LOCK:
    ensure_custom_local_apps_file()

with USER_PROFILE_LOCK:
    ensure_user_profile_file()

USER_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
ensure_brand_image_file()


# ============================================================
# FRONTEND - inspirado no dashboard de referência enviado
# ============================================================

INDEX_HTML = r'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        hub: {
                            950: '#06121d',
                            900: '#081824',
                            850: '#0a1d2b',
                            800: '#0c2232',
                            750: '#10293a',
                            cyan: '#35d5e6'
                        }
                    },
                    boxShadow: {
                        panel: '0 24px 70px rgba(0,0,0,.32)',
                        cyan: '0 0 24px rgba(53,213,230,.12)'
                    }
                }
            }
        }
    </script>

    <style>
        :root {
            --bg: #06121d;
            --panel: #0a1d2b;
            --panel-2: #0c2232;
            --line: rgba(129, 180, 204, .14);
            --text-muted: #7f9caf;
            --cyan: #35d5e6;
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }
        html { scroll-behavior: smooth; }

        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 76% 5%, rgba(44, 196, 219, .10), transparent 27rem),
                radial-gradient(circle at 9% 82%, rgba(19, 96, 129, .11), transparent 26rem),
                linear-gradient(145deg, #05101a 0%, #071722 46%, #05111c 100%);
        }

        body::before {
            content: '';
            position: fixed;
            inset: 0;
            pointer-events: none;
            opacity: .12;
            background-image:
                linear-gradient(rgba(76, 188, 216, .06) 1px, transparent 1px),
                linear-gradient(90deg, rgba(76, 188, 216, .05) 1px, transparent 1px);
            background-size: 44px 44px;
            mask-image: linear-gradient(to bottom, black, transparent 86%);
        }

        .hub-panel {
            background: linear-gradient(145deg, rgba(13, 35, 50, .94), rgba(8, 25, 37, .96));
            border: 1px solid var(--line);
            box-shadow: inset 0 1px rgba(255,255,255,.018);
        }

        .glass-top {
            background: rgba(7, 24, 36, .88);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .quick-card, .tool-card, .sidebar-item, .category-chip {
            transition: transform .18s ease, border-color .18s ease, background .18s ease, box-shadow .18s ease;
        }

        .quick-card:hover, .tool-card:hover {
            transform: translateY(-3px);
            border-color: rgba(53, 213, 230, .38);
            box-shadow: 0 14px 35px rgba(0,0,0,.22), 0 0 20px rgba(53,213,230,.045);
        }

        .sidebar-item:hover, .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28, 141, 169, .27), rgba(14, 72, 94, .13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
            box-shadow: 0 0 14px rgba(53,213,230,.65);
        }

        .category-chip.active {
            color: #dffaff;
            border-color: rgba(53,213,230,.45);
            background: rgba(33, 152, 180, .15);
        }

        .truncate-2 {
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }

        .modal-backdrop {
            background: rgba(1, 8, 13, .78);
            backdrop-filter: blur(7px);
            -webkit-backdrop-filter: blur(7px);
        }

        .fade-in { animation: fadeIn .2s ease-out; }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(8px) scale(.99); }
            to { opacity: 1; transform: translateY(0) scale(1); }
        }

        @media (max-width: 1023px) {
            #sidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }
            #sidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased selection:bg-cyan-400/20 selection:text-cyan-100">

    <!-- OVERLAY MOBILE -->
    <div id="sidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <!-- SIDEBAR -->
    <aside id="sidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <div class="flex h-[132px] items-center border-b border-slate-700/20 px-3 py-3">
            <a href="/" class="flex w-full flex-col items-center justify-center text-center">
                <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain" onerror="this.style.display='none'; this.nextElementSibling.classList.remove('hidden');">
                <div class="hidden items-center justify-center rounded-2xl border border-cyan-400/20 bg-cyan-400/10 p-4 text-3xl shadow-cyan">⬡</div>
                <div class="mt-2 leading-tight">
                    <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                    <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
                </div>
            </a>
        </div>

        <div class="border-b border-slate-700/20 px-5 py-4">
            <div id="sidebarDate" class="text-xs font-semibold text-slate-300">—</div>
            <div id="sidebarTime" class="mt-1 text-[11px] text-slate-500">—</div>
        </div>

        <nav class="flex-1 overflow-y-auto px-3 py-4">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">▦</span>
                <span class="flex-1 font-semibold">APPS + USADOS</span>
            </a>

            <a href="/local-apps" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">▣</span>
                <span class="flex-1 font-semibold">APPS LOCAIS</span>
            </a>

            <a href="/links" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">★</span>
                <span class="flex-1 font-semibold">LINKS FAVORITOS</span>
            </a>

            <a href="/workspace" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">◈</span>
                <span class="flex-1 font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>

            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span>
                <span class="flex-1 font-semibold">DIAGNÓSTICO WIN</span>
            </a>

            <button data-scroll-tools class="sidebar-item active relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400">
                <span class="w-5 text-center">☷</span>
                <span class="flex-1 font-semibold">All tools</span>
                <span id="sidebarTotal" class="text-[10px] text-slate-600">0</span>
            </button>

            <p class="mb-2 mt-6 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Categorias</p>
            <div id="sidebarCategories" class="space-y-1"></div>
        </nav>

        <div class="border-t border-slate-700/20 p-3">
            <a href="/docs" target="_blank" rel="noopener noreferrer" class="sidebar-item relative flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⚡</span>
                <span>Documentação API</span>
            </a>
            <div class="mt-3 rounded-xl border border-slate-700/20 bg-[#081b28] p-3">
                <div class="flex items-center gap-2">
                    <span class="inline-block h-2 w-2 rounded-full bg-emerald-400 shadow-[0_0_10px_rgba(52,211,153,.6)]"></span>
                    <span class="text-xs font-semibold text-slate-300">Sistema operacional</span>
                </div>
                <p class="mt-1 pl-4 text-[10px] text-slate-600">FastAPI + JSON local</p>
            </div>
        </div>
    </aside>

    <!-- ÁREA PRINCIPAL -->
    <div class="min-h-screen lg:pl-[242px]">

        <!-- TOPBAR -->
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex h-[72px] items-center gap-3 px-4 sm:px-6 xl:px-8">
                <button id="mobileMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden" aria-label="Abrir menu">☰</button>

                <div class="relative min-w-0 flex-1 max-w-3xl">
                    <span class="pointer-events-none absolute left-4 top-1/2 -translate-y-1/2 text-slate-500">⌕</span>
                    <input
                        id="searchInput"
                        type="search"
                        placeholder="Buscar ferramentas, categorias ou descrições..."
                        autocomplete="off"
                        class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#0b2232]/85 pl-11 pr-4 text-sm text-slate-200 outline-none transition placeholder:text-slate-600 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10"
                    >
                </div>

                <a href="/docs" target="_blank" rel="noopener noreferrer" class="hidden rounded-lg px-3 py-2 text-xs font-medium text-slate-400 transition hover:bg-slate-800/40 hover:text-cyan-300 md:block">API</a>

                <button id="openModalBtn" type="button" class="inline-flex h-11 items-center gap-2 rounded-xl border border-cyan-300/20 bg-cyan-400/10 px-4 text-sm font-bold text-cyan-200 shadow-cyan transition hover:border-cyan-300/40 hover:bg-cyan-400/15">
                    <span class="text-lg leading-none">＋</span>
                    <span class="hidden sm:inline">Adicionar App</span>
                </button>

                <button id="profileUploadTrigger" type="button" class="hidden items-center gap-3 border-l border-slate-700/25 pl-4 text-left xl:flex" title="Clique para alterar a imagem do perfil">
                    <div class="grid h-9 w-9 place-items-center overflow-hidden rounded-full border border-slate-700/30 bg-[#0c2637]">
                        <img id="profileAvatar" src="" alt="Foto de perfil" class="hidden h-full w-full object-cover">
                        <span id="profileAvatarFallback">👤</span>
                    </div>
                    <div class="leading-tight">
                        <div id="profileName" class="text-xs font-semibold text-slate-300">Usuário local</div>
                        <div id="profileRole" class="text-[10px] text-slate-600">Administrador</div>
                    </div>
                    <input id="profileImageInput" type="file" accept="image/png,image/jpeg,image/webp,image/gif" class="hidden">
                </button>
            </div>
        </header>

        <main class="mx-auto max-w-[1550px] px-4 py-5 sm:px-6 xl:px-8">

            <!-- CABEÇALHO DA PÁGINA -->
            <section class="mb-5 flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
                <div>
                    <p class="mb-1 text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Dashboard</p>
                    <h1 class="text-2xl font-bold tracking-tight text-slate-100 sm:text-3xl">Central de Ferramentas</h1>
                    <p class="mt-1 text-sm text-slate-500">Acesse e gerencie seus aplicativos em um único painel.</p>
                </div>

                <div class="flex flex-wrap items-center justify-end gap-2 text-xs text-slate-500">
                    <span class="inline-flex items-center gap-1.5 rounded-lg border border-slate-700/25 bg-[#091c29] px-3 py-2">
                        <span class="h-1.5 w-1.5 rounded-full bg-emerald-400"></span>
                        API online
                    </span>
                </div>
            </section>

            <!-- ACESSO RÁPIDO -->
            <section class="hub-panel mb-5 rounded-2xl p-4 sm:p-5">
                <div class="mb-4 flex items-center justify-between gap-3">
                    <div>
                        <h2 class="text-sm font-bold text-slate-200">Acesso rápido</h2>
                        <p class="mt-0.5 text-[11px] text-slate-600">Abra suas principais ferramentas em um clique</p>
                    </div>
                    <span id="quickCount" class="rounded-md border border-cyan-400/10 bg-cyan-400/5 px-2 py-1 text-[10px] font-semibold text-cyan-400/70">0 apps</span>
                </div>

                <div id="quickAccessGrid" class="grid grid-cols-2 gap-2.5 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-6 2xl:grid-cols-7"></div>
            </section>

            <div class="grid gap-5 xl:grid-cols-[minmax(0,1fr)_320px]">

                <!-- COLUNA PRINCIPAL -->
                <div class="min-w-0 space-y-5">
                    <section id="toolsSection" class="hub-panel rounded-2xl p-4 sm:p-5">
                        <div class="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                            <div>
                                <h2 class="text-sm font-bold text-slate-200">Ferramentas cadastradas</h2>
                                <p id="resultsLabel" class="mt-0.5 text-[11px] text-slate-600">Carregando...</p>
                            </div>

                            <div id="categoryChips" class="flex max-w-full gap-2 overflow-x-auto pb-1"></div>
                        </div>

                        <div id="loadingState" class="py-16 text-center text-sm text-slate-600">Carregando aplicativos...</div>
                        <div id="appsGrid" class="grid grid-cols-1 gap-3 md:grid-cols-2 2xl:grid-cols-3"></div>

                        <div id="emptyState" class="hidden rounded-xl border border-dashed border-slate-700/30 bg-[#071924]/70 px-5 py-14 text-center">
                            <div class="text-4xl">⌕</div>
                            <h3 class="mt-3 text-sm font-bold text-slate-300">Nenhum aplicativo encontrado</h3>
                            <p class="mt-1 text-xs text-slate-600">Altere o filtro, a busca ou cadastre uma nova ferramenta.</p>
                        </div>
                    </section>
                </div>

                <!-- COLUNA DE STATUS -->
                <aside class="space-y-5">
                    <section class="hub-panel rounded-2xl p-4">
                        <div class="mb-4 flex items-center justify-between">
                            <h2 class="text-sm font-bold text-slate-200">Status do sistema</h2>
                            <span class="text-[10px] text-slate-600">Visão geral</span>
                        </div>

                        <div class="grid grid-cols-2 gap-2.5">
                            <div class="rounded-xl border border-slate-700/20 bg-[#071a27] p-3">
                                <div class="text-[10px] uppercase tracking-[.12em] text-slate-600">Apps</div>
                                <div id="appsCount" class="mt-1 text-2xl font-bold text-cyan-300">0</div>
                            </div>
                            <div class="rounded-xl border border-slate-700/20 bg-[#071a27] p-3">
                                <div class="text-[10px] uppercase tracking-[.12em] text-slate-600">Categorias</div>
                                <div id="categoriesCount" class="mt-1 text-2xl font-bold text-emerald-300">0</div>
                            </div>
                        </div>

                        <div class="mt-3 space-y-2">
                            <div class="flex items-center justify-between rounded-lg border border-slate-700/15 bg-[#071a27]/75 px-3 py-2.5">
                                <div class="flex items-center gap-2">
                                    <span class="h-2 w-2 rounded-full bg-emerald-400"></span>
                                    <span class="text-xs text-slate-400">FastAPI</span>
                                </div>
                                <span class="text-[10px] font-semibold text-emerald-400">Operacional</span>
                            </div>
                            <div class="flex items-center justify-between rounded-lg border border-slate-700/15 bg-[#071a27]/75 px-3 py-2.5">
                                <div class="flex items-center gap-2">
                                    <span class="h-2 w-2 rounded-full bg-emerald-400"></span>
                                    <span class="text-xs text-slate-400">Armazenamento JSON</span>
                                </div>
                                <span class="text-[10px] font-semibold text-emerald-400">Disponível</span>
                            </div>
                        </div>
                    </section>

                    <section class="hub-panel rounded-2xl p-4">
                        <div class="mb-4 flex items-center justify-between">
                            <h2 class="text-sm font-bold text-slate-200">Categorias</h2>
                            <span class="text-[10px] text-slate-600">Distribuição</span>
                        </div>
                        <div id="categoryStats" class="space-y-3"></div>
                    </section>

                    <section class="hub-panel rounded-2xl p-4">
                        <h2 class="text-sm font-bold text-slate-200">Atalhos</h2>
                        <div class="mt-3 grid gap-2">
                            <button id="sideAddBtn" type="button" class="flex items-center justify-between rounded-lg border border-slate-700/20 bg-[#071a27] px-3 py-2.5 text-left text-xs text-slate-400 transition hover:border-cyan-400/25 hover:text-cyan-200">
                                <span>Adicionar ferramenta</span><span>＋</span>
                            </button>
                            <a href="/docs" target="_blank" rel="noopener noreferrer" class="flex items-center justify-between rounded-lg border border-slate-700/20 bg-[#071a27] px-3 py-2.5 text-xs text-slate-400 transition hover:border-cyan-400/25 hover:text-cyan-200">
                                <span>Abrir Swagger API</span><span>↗</span>
                            </a>
                            <button id="clearFiltersBtn" type="button" class="flex items-center justify-between rounded-lg border border-slate-700/20 bg-[#071a27] px-3 py-2.5 text-left text-xs text-slate-400 transition hover:border-cyan-400/25 hover:text-cyan-200">
                                <span>Limpar filtros</span><span>↺</span>
                            </button>
                        </div>
                    </section>
                </aside>
            </div>

            <footer class="mt-6 border-t border-slate-700/15 py-5 text-center text-[10px] uppercase tracking-[.15em] text-slate-700">
                TECH TOOL HUB • FastAPI • Tailwind CSS • JSON Local
            </footer>
        </main>
    </div>

    <!-- MODAL CADASTRO -->
    <div id="appModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true" aria-labelledby="modalTitle">
        <div class="fade-in w-full max-w-xl overflow-hidden rounded-2xl border border-cyan-400/15 bg-[#081925] shadow-[0_28px_90px_rgba(0,0,0,.55)]">
            <div class="flex items-start justify-between border-b border-slate-700/20 bg-[#0a1f2d] px-5 py-4 sm:px-6">
                <div>
                    <p class="text-[10px] font-bold uppercase tracking-[.16em] text-cyan-400/60">TECH TOOL HUB</p>
                    <h2 id="modalTitle" class="mt-1 text-xl font-bold text-slate-100">Adicionar aplicativo</h2>
                    <p id="modalSubtitle" class="mt-1 text-xs text-slate-500">Cadastre uma ferramenta para acesso rápido.</p>
                </div>
                <button id="closeModalX" type="button" class="grid h-9 w-9 place-items-center rounded-lg text-slate-500 transition hover:bg-slate-800/50 hover:text-slate-200" aria-label="Fechar modal">✕</button>
            </div>

            <form id="appForm" class="space-y-4 p-5 sm:p-6">
                <div class="grid gap-4 sm:grid-cols-[105px_1fr]">
                    <div>
                        <label for="appId" class="mb-1.5 block text-xs font-semibold text-slate-400">ID</label>
                        <input id="appId" name="id" type="number" min="1" required readonly class="h-11 w-full rounded-lg border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-500 outline-none">
                    </div>
                    <div>
                        <label for="appName" class="mb-1.5 block text-xs font-semibold text-slate-400">Nome</label>
                        <input id="appName" name="name" type="text" maxlength="80" required placeholder="Ex.: VS Code" class="h-11 w-full rounded-lg border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                    </div>
                </div>

                <div class="grid gap-4 sm:grid-cols-[180px_1fr]">
                    <div>
                        <label for="appIcon" class="mb-1.5 block text-xs font-semibold text-slate-400">Ícone do App</label>
                        <label class="flex h-[116px] cursor-pointer flex-col items-center justify-center rounded-xl border border-dashed border-slate-700/40 bg-[#06131d] px-3 text-center transition hover:border-cyan-400/35 hover:bg-cyan-400/[.03]">
                            <div id="iconPreview" class="mb-2 grid h-12 w-12 place-items-center overflow-hidden rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-2xl text-slate-600">＋</div>
                            <span id="iconFileLabel" class="max-w-full truncate text-[11px] font-semibold text-slate-500">Selecionar imagem</span>
                            <span class="mt-1 text-[9px] text-slate-700">PNG, JPG, WEBP ou GIF • até 2 MB</span>
                            <input id="appIcon" name="icon_file" type="file" accept="image/png,image/jpeg,image/webp,image/gif" class="sr-only">
                        </label>
                    </div>
                    <div>
                        <label for="appCategory" class="mb-1.5 block text-xs font-semibold text-slate-400">Categoria</label>
                        <input id="appCategory" name="category" type="text" maxlength="50" required placeholder="Ex.: Dev, IA, Produtividade" class="h-11 w-full rounded-lg border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                        <p class="mt-2 text-[10px] leading-4 text-slate-700">O arquivo do ícone será armazenado localmente na pasta <strong class="font-semibold text-slate-600">icons</strong>.</p>
                    </div>
                </div>

                <div>
                    <label for="appUrl" class="mb-1.5 block text-xs font-semibold text-slate-400">URL</label>
                    <input id="appUrl" name="url" type="url" maxlength="500" required placeholder="https://..." class="h-11 w-full rounded-lg border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                </div>

                <div>
                    <label for="appDescription" class="mb-1.5 block text-xs font-semibold text-slate-400">Descrição</label>
                    <textarea id="appDescription" name="description" rows="3" maxlength="280" required placeholder="Descreva rapidamente a utilidade da ferramenta..." class="w-full resize-none rounded-lg border border-slate-700/30 bg-[#06131d] px-3 py-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10"></textarea>
                </div>

                <div id="formError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:justify-end">
                    <button id="cancelBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 transition hover:bg-slate-800/35 hover:text-slate-200">Cancelar</button>
                    <button id="saveBtn" type="submit" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 transition hover:border-cyan-300/35 hover:bg-cyan-400/15 disabled:cursor-not-allowed disabled:opacity-50">Salvar aplicativo</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL CRIAR EXECUTÁVEL -->
    <div id="buildExecutableModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-3xl overflow-hidden rounded-2xl border border-slate-700/30 bg-[#091d2a] shadow-2xl">
            <div class="flex items-start justify-between gap-4 border-b border-slate-700/20 px-5 py-4">
                <div>
                    <h3 class="text-lg font-bold text-slate-100">Criar Executável</h3>
                    <p class="mt-1 text-xs text-slate-500">Gera TechToolHub.exe com PyInstaller e mostra o log em tempo real.</p>
                </div>
                <button id="closeBuildExecutableModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <div class="space-y-4 p-5">
                <div id="buildEnvironmentBox" class="rounded-xl border border-slate-700/25 bg-[#061722] p-4 text-xs text-slate-500">Verificando ambiente...</div>

                <div class="grid gap-3 sm:grid-cols-3">
                    <div class="rounded-xl border border-slate-700/20 bg-[#071a27] p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Status</div>
                        <div id="buildStatusLabel" class="mt-1 text-sm font-bold text-slate-300">Pronto</div>
                    </div>
                    <div class="rounded-xl border border-slate-700/20 bg-[#071a27] p-3 sm:col-span-2">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Saída</div>
                        <div id="buildOutputPath" class="mt-1 truncate text-xs font-semibold text-slate-400">—</div>
                    </div>
                </div>

                <div>
                    <div class="mb-2 flex items-center justify-between gap-3">
                        <span class="text-xs font-bold text-slate-300">Log do build</span>
                        <span id="buildRunningBadge" class="hidden rounded-md border border-cyan-400/15 bg-cyan-400/5 px-2 py-1 text-[9px] font-bold uppercase tracking-[.08em] text-cyan-300">Processando</span>
                    </div>
                    <pre id="buildLog" class="max-h-[320px] min-h-[170px] overflow-auto whitespace-pre-wrap rounded-xl border border-slate-700/20 bg-[#04111a] p-4 text-[11px] leading-5 text-slate-500">Aguardando início do build...</pre>
                </div>

                <div id="buildError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 sm:flex-row sm:justify-between">
                    <button id="openBuildFolderBtn" type="button" disabled class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 disabled:cursor-not-allowed disabled:opacity-35">Abrir pasta de saída</button>
                    <div class="flex flex-col-reverse gap-2 sm:flex-row">
                        <button id="cancelBuildModalBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400">Fechar</button>
                        <button id="runExecutableBuildBtn" type="button" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 disabled:cursor-not-allowed disabled:opacity-40">Gerar EXE</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- MODAL README -->
    <div id="readmeModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-4xl overflow-hidden rounded-2xl border border-slate-700/30 bg-[#091d2a] shadow-2xl">
            <div class="flex items-start justify-between gap-4 border-b border-slate-700/20 px-5 py-4">
                <div>
                    <h3 class="text-lg font-bold text-slate-100">README — TECH TOOL HUB</h3>
                    <p class="mt-1 text-xs text-slate-500">Execução, dados, geração do EXE e MSIX.</p>
                </div>
                <button id="closeReadmeModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>
            <div class="p-5">
                <pre id="readmeContent" class="max-h-[65vh] overflow-auto whitespace-pre-wrap rounded-xl border border-slate-700/20 bg-[#04111a] p-4 text-xs leading-6 text-slate-400">Carregando README...</pre>
            </div>
        </div>
    </div>

    <!-- TOAST -->
    <div id="toast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const state = {
            apps: [],
            search: '',
            category: 'all',
            editAppId: null,
            currentIcon: ''
        };

        const appsGrid = document.getElementById('appsGrid');
        const quickAccessGrid = document.getElementById('quickAccessGrid');
        const emptyState = document.getElementById('emptyState');
        const loadingState = document.getElementById('loadingState');
        const appsCount = document.getElementById('appsCount');
        const categoriesCount = document.getElementById('categoriesCount');
        const sidebarTotal = document.getElementById('sidebarTotal');
        const quickCount = document.getElementById('quickCount');
        const resultsLabel = document.getElementById('resultsLabel');
        const categoryStats = document.getElementById('categoryStats');
        const sidebarCategories = document.getElementById('sidebarCategories');
        const categoryChips = document.getElementById('categoryChips');
        const searchInput = document.getElementById('searchInput');

        const modal = document.getElementById('appModal');
        const form = document.getElementById('appForm');
        const modalTitle = document.getElementById('modalTitle');
        const modalSubtitle = document.getElementById('modalSubtitle');
        const formError = document.getElementById('formError');
        const saveBtn = document.getElementById('saveBtn');
        const appId = document.getElementById('appId');
        const appName = document.getElementById('appName');
        const appCategory = document.getElementById('appCategory');
        const appUrl = document.getElementById('appUrl');
        const appIcon = document.getElementById('appIcon');
        const iconPreview = document.getElementById('iconPreview');
        const iconFileLabel = document.getElementById('iconFileLabel');
        const appDescription = document.getElementById('appDescription');
        const toast = document.getElementById('toast');

        const profileUploadTrigger = document.getElementById('profileUploadTrigger');
        const profileImageInput = document.getElementById('profileImageInput');
        const profileAvatar = document.getElementById('profileAvatar');
        const profileAvatarFallback = document.getElementById('profileAvatarFallback');
        const profileName = document.getElementById('profileName');
        const profileRole = document.getElementById('profileRole');

        const createExecutableBtn = document.getElementById('createExecutableBtn');
        const showReadmeBtn = document.getElementById('showReadmeBtn');
        const buildExecutableModal = document.getElementById('buildExecutableModal');
        const buildEnvironmentBox = document.getElementById('buildEnvironmentBox');
        const buildStatusLabel = document.getElementById('buildStatusLabel');
        const buildOutputPath = document.getElementById('buildOutputPath');
        const buildLog = document.getElementById('buildLog');
        const buildRunningBadge = document.getElementById('buildRunningBadge');
        const buildError = document.getElementById('buildError');
        const runExecutableBuildBtn = document.getElementById('runExecutableBuildBtn');
        const openBuildFolderBtn = document.getElementById('openBuildFolderBtn');
        const readmeModal = document.getElementById('readmeModal');
        const readmeContent = document.getElementById('readmeContent');
        let buildPollTimer = null;

        const sidebar = document.getElementById('sidebar');
        const sidebarOverlay = document.getElementById('sidebarOverlay');
        const mobileMenuBtn = document.getElementById('mobileMenuBtn');

        function setModalVisible(modal, visible) {
            modal.classList.toggle('hidden', !visible);
            modal.classList.toggle('flex', visible);
        }

        function buildStatusText(status) {
            if (status === 'starting') return 'Preparando...';
            if (status === 'running') return 'Gerando...';
            if (status === 'success') return 'Concluído';
            if (status === 'error') return 'Falhou';
            return 'Pronto';
        }

        function stopBuildPolling() {
            if (!buildPollTimer) return;
            clearInterval(buildPollTimer);
            buildPollTimer = null;
        }

        function startBuildPolling() {
            if (buildPollTimer) return;
            buildPollTimer = setInterval(refreshBuildStatus, 1000);
        }

        function renderBuildState(data) {
            const environment = data.environment || {};
            const reasons = Array.isArray(environment.reasons) ? environment.reasons : [];

            buildEnvironmentBox.innerHTML = environment.available
                ? `<div class="flex items-start gap-3">
                       <span class="mt-1 h-2 w-2 shrink-0 rounded-full bg-emerald-400"></span>
                       <div>
                           <div class="font-semibold text-emerald-300">Ambiente pronto para gerar o EXE.</div>
                           <div class="mt-1 break-all text-[10px] text-slate-600">${escapeHtml(environment.source_dir || '')}</div>
                       </div>
                   </div>`
                : `<div class="flex items-start gap-3">
                       <span class="mt-1 h-2 w-2 shrink-0 rounded-full bg-amber-300"></span>
                       <div>
                           <div class="font-semibold text-amber-200">Build indisponível neste modo.</div>
                           <div class="mt-1 space-y-1 text-[10px] text-slate-600">
                               ${reasons.map(reason => `<div>• ${escapeHtml(reason)}</div>`).join('')}
                           </div>
                       </div>
                   </div>`;

            buildStatusLabel.textContent = buildStatusText(data.status);
            buildOutputPath.textContent = data.output_path || environment.output_dir || '—';
            buildLog.textContent = (data.log || []).join('\n') || data.message || 'Aguardando início do build...';

            const running = Boolean(data.running);
            buildRunningBadge.classList.toggle('hidden', !running);
            runExecutableBuildBtn.disabled = running || !environment.available;
            runExecutableBuildBtn.textContent = running ? 'Gerando...' : 'Gerar EXE';
            openBuildFolderBtn.disabled = data.status !== 'success';

            if (data.status === 'error') {
                buildError.textContent = data.message || 'O build falhou.';
                buildError.classList.remove('hidden');
            } else {
                buildError.classList.add('hidden');
                buildError.textContent = '';
            }

            if (running) startBuildPolling();
            else stopBuildPolling();
        }

        async function refreshBuildStatus() {
            try {
                const response = await fetch('/api/build/status', { cache: 'no-store' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível consultar o build.');
                renderBuildState(data);
            } catch (error) {
                buildError.textContent = error.message || 'Erro ao consultar o build.';
                buildError.classList.remove('hidden');
            }
        }

        async function openBuildExecutableModal() {
            setModalVisible(buildExecutableModal, true);
            await refreshBuildStatus();
        }

        function closeBuildExecutableModal() {
            setModalVisible(buildExecutableModal, false);
            stopBuildPolling();
        }

        async function runExecutableBuild() {
            runExecutableBuildBtn.disabled = true;
            buildError.classList.add('hidden');

            try {
                const response = await fetch('/api/build/executable', { method: 'POST' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível iniciar o build.');
                renderBuildState(data);
                startBuildPolling();
            } catch (error) {
                buildError.textContent = error.message || 'Erro ao iniciar o build.';
                buildError.classList.remove('hidden');
                await refreshBuildStatus();
            }
        }

        async function openBuildOutputFolder() {
            try {
                const response = await fetch('/api/build/open-output', { method: 'POST' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível abrir a pasta.');
                showToast(data.message || 'Pasta aberta.');
            } catch (error) {
                showToast(error.message || 'Erro ao abrir a pasta.', true);
            }
        }

        async function openReadmeModal() {
            setModalVisible(readmeModal, true);
            readmeContent.textContent = 'Carregando README...';

            try {
                const response = await fetch('/api/readme', { cache: 'no-store' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível carregar o README.');
                readmeContent.textContent = data.content || 'README vazio.';
            } catch (error) {
                readmeContent.textContent = `Erro ao carregar README:\n${error.message || error}`;
            }
        }

        function closeReadmeModal() {
            setModalVisible(readmeModal, false);
        }

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function iconMarkup(icon, imageClass = 'h-full w-full object-contain') {
            const value = String(icon ?? '').trim();
            const isImage = value.startsWith('/icons/') || /^https?:\/\//i.test(value);
            if (isImage) {
                return `<img src="${escapeHtml(value)}" alt="" class="${imageClass}" loading="lazy" referrerpolicy="no-referrer" onerror="this.style.display='none';this.parentElement.textContent='🔗';">`;
            }
            return escapeHtml(value || '🔗');
        }

        function safeUrl(value) {
            try {
                const url = new URL(value);
                return ['http:', 'https:'].includes(url.protocol) ? url.href : '#';
            } catch {
                return '#';
            }
        }

        function normalize(value) {
            return String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase();
        }

        function applyUserProfile(profile) {
            profileName.textContent = profile?.name || 'Usuário local';
            profileRole.textContent = profile?.role || 'Administrador';

            const imageUrl = String(profile?.image || '').trim();
            if (imageUrl) {
                profileAvatar.src = `${imageUrl}${imageUrl.includes('?') ? '&' : '?'}v=${Date.now()}`;
                profileAvatar.classList.remove('hidden');
                profileAvatarFallback.classList.add('hidden');
            } else {
                profileAvatar.src = '';
                profileAvatar.classList.add('hidden');
                profileAvatarFallback.classList.remove('hidden');
            }
        }

        async function loadUserProfile() {
            try {
                const response = await fetch('/api/user-profile', { cache: 'no-store' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível carregar o perfil.');
                applyUserProfile(data);
            } catch {
                applyUserProfile({ name: 'Usuário local', role: 'Administrador', image: '' });
            }
        }

        async function uploadUserProfileImage(file) {
            if (!file) return;
            if (!file.type.startsWith('image/')) {
                showToast('Selecione uma imagem válida para o perfil.', 'error');
                return;
            }
            if (file.size > 5 * 1024 * 1024) {
                showToast('A imagem do perfil deve ter no máximo 5 MB.', 'error');
                return;
            }

            try {
                const response = await fetch('/api/user-profile/image', {
                    method: 'POST',
                    headers: { 'Content-Type': file.type || 'application/octet-stream' },
                    body: file
                });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Não foi possível enviar a imagem do perfil.');
                applyUserProfile(data);
                showToast('Imagem de perfil atualizada.');
            } catch (error) {
                showToast(error.message || 'Erro ao atualizar o perfil.', 'error');
            } finally {
                profileImageInput.value = '';
            }
        }

        function updateClock() {
            const now = new Date();
            const dateText = new Intl.DateTimeFormat('pt-BR', {
                weekday: 'short', day: '2-digit', month: 'short', year: 'numeric'
            }).format(now);
            const timeText = new Intl.DateTimeFormat('pt-BR', {
                hour: '2-digit', minute: '2-digit'
            }).format(now);

            document.getElementById('sidebarDate').textContent = dateText;
            document.getElementById('sidebarTime').textContent = `${timeText} • ambiente local`;
        }

        function nextSuggestedId() {
            if (!state.apps.length) return 1;
            return Math.max(...state.apps.map(app => Number(app.id) || 0)) + 1;
        }

        function showToast(message, type = 'success') {
            toast.textContent = message;
            toast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (type === 'error'
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            toast.classList.remove('hidden');
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => toast.classList.add('hidden'), 2800);
        }

        function getCategoryMap() {
            const map = new Map();
            state.apps.forEach(app => {
                const category = String(app.category || 'Outros').trim() || 'Outros';
                map.set(category, (map.get(category) || 0) + 1);
            });
            return [...map.entries()].sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));
        }

        function filteredApps() {
            const query = normalize(state.search.trim());
            const category = normalize(state.category);

            return [...state.apps]
                .filter(app => {
                    const categoryMatch = state.category === 'all' || normalize(app.category) === category;
                    if (!categoryMatch) return false;

                    if (!query) return true;
                    return [app.name, app.category, app.description]
                        .some(value => normalize(value).includes(query));
                })
                .sort((a, b) => String(a.name).localeCompare(String(b.name), 'pt-BR'));
        }

        function renderCategories() {
            const categories = getCategoryMap();
            categoriesCount.textContent = categories.length;

            sidebarCategories.innerHTML = categories.map(([category, count]) => `
                <button data-category="${escapeHtml(category)}" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-left text-sm text-slate-400 ${state.category === category ? 'active' : ''}">
                    <span class="w-5 text-center text-cyan-400/55">◇</span>
                    <span class="min-w-0 flex-1 truncate">${escapeHtml(category)}</span>
                    <span class="text-[10px] text-slate-600">${count}</span>
                </button>
            `).join('');

            categoryChips.innerHTML = `
                <button data-chip-category="all" class="category-chip shrink-0 rounded-lg border border-slate-700/25 bg-[#071a27] px-3 py-1.5 text-[11px] font-semibold text-slate-500 ${state.category === 'all' ? 'active' : ''}">Todos</button>
                ${categories.map(([category]) => `
                    <button data-chip-category="${escapeHtml(category)}" class="category-chip shrink-0 rounded-lg border border-slate-700/25 bg-[#071a27] px-3 py-1.5 text-[11px] font-semibold text-slate-500 ${state.category === category ? 'active' : ''}">${escapeHtml(category)}</button>
                `).join('')}
            `;

            const max = Math.max(1, ...categories.map(([, count]) => count));
            categoryStats.innerHTML = categories.length
                ? categories.map(([category, count]) => {
                    const width = Math.max(10, Math.round((count / max) * 100));
                    return `
                        <button data-stat-category="${escapeHtml(category)}" class="block w-full text-left">
                            <div class="mb-1.5 flex items-center justify-between text-[11px]">
                                <span class="truncate pr-3 text-slate-400">${escapeHtml(category)}</span>
                                <span class="font-semibold text-slate-600">${count}</span>
                            </div>
                            <div class="h-1.5 overflow-hidden rounded-full bg-[#06131d]">
                                <div class="h-full rounded-full bg-gradient-to-r from-cyan-500/55 to-cyan-300/80" style="width:${width}%"></div>
                            </div>
                        </button>
                    `;
                }).join('')
                : '<p class="text-xs text-slate-600">Nenhuma categoria cadastrada.</p>';
        }

        function renderQuickAccess() {
            const visible = filteredApps().slice(0, 7);
            quickCount.textContent = `${visible.length} ${visible.length === 1 ? 'app' : 'apps'}`;

            if (!visible.length) {
                quickAccessGrid.innerHTML = '<div class="col-span-full py-5 text-center text-xs text-slate-600">Nenhuma ferramenta para acesso rápido.</div>';
                return;
            }

            quickAccessGrid.innerHTML = visible.map(app => `
                <a href="${escapeHtml(safeUrl(app.url))}" target="_blank" rel="noopener noreferrer" class="quick-card group min-w-0 rounded-xl border border-slate-700/25 bg-[#091f2e] p-3">
                    <div class="flex items-start justify-between gap-2">
                        <div class="grid h-9 w-9 shrink-0 place-items-center overflow-hidden rounded-lg border border-slate-700/25 bg-[#061722] text-xl">${iconMarkup(app.icon, 'h-7 w-7 object-contain')}</div>
                        <span class="mt-1 h-1.5 w-1.5 rounded-full bg-emerald-400/80 shadow-[0_0_8px_rgba(52,211,153,.45)]"></span>
                    </div>
                    <div class="mt-3 truncate text-xs font-bold text-slate-300 transition group-hover:text-cyan-200">${escapeHtml(app.name)}</div>
                    <div class="mt-1 truncate text-[10px] text-slate-600">${escapeHtml(app.category)}</div>
                </a>
            `).join('');
        }

        function renderApps() {
            const visible = filteredApps();

            appsCount.textContent = state.apps.length;
            sidebarTotal.textContent = state.apps.length;
            resultsLabel.textContent = `${visible.length} de ${state.apps.length} ${state.apps.length === 1 ? 'ferramenta' : 'ferramentas'}`;

            appsGrid.innerHTML = '';
            emptyState.classList.toggle('hidden', visible.length !== 0);

            visible.forEach(app => {
                const card = document.createElement('article');
                card.className = 'tool-card group flex min-h-[178px] flex-col rounded-xl border border-slate-700/25 bg-[#091d2a] p-4';

                const escapedName = escapeHtml(app.name);
                const escapedCategory = escapeHtml(app.category);
                const escapedDescription = escapeHtml(app.description);
                const renderedIcon = iconMarkup(app.icon, 'h-9 w-9 object-contain');
                const escapedUrl = escapeHtml(safeUrl(app.url));

                card.innerHTML = `
                    <div class="flex items-start gap-3">
                        <div class="grid h-11 w-11 shrink-0 place-items-center overflow-hidden rounded-xl border border-slate-700/25 bg-[#061722] text-2xl shadow-inner">${renderedIcon}</div>
                        <div class="min-w-0 flex-1">
                            <div class="flex items-start justify-between gap-2">
                                <div class="min-w-0">
                                    <h3 class="truncate text-sm font-bold text-slate-200 transition group-hover:text-cyan-100">${escapedName}</h3>
                                    <span class="mt-1 inline-flex rounded-md border border-cyan-400/10 bg-cyan-400/5 px-2 py-0.5 text-[9px] font-bold uppercase tracking-[.08em] text-cyan-400/65">${escapedCategory}</span>
                                </div>
                                <span class="mt-1 h-1.5 w-1.5 shrink-0 rounded-full bg-emerald-400/80"></span>
                            </div>
                        </div>
                    </div>

                    <p class="truncate-2 mt-3 flex-1 text-xs leading-5 text-slate-500">${escapedDescription}</p>

                    <div class="mt-4 flex items-center justify-between gap-3 border-t border-slate-700/15 pt-3">
                        <a href="${escapedUrl}" target="_blank" rel="noopener noreferrer" class="text-[11px] font-bold text-cyan-400/75 transition hover:text-cyan-200">Acessar App ↗</a>
                        <div class="flex items-center gap-1">
                            <button type="button" data-edit-id="${Number(app.id)}" class="rounded-md px-2 py-1 text-[10px] text-slate-500 transition hover:bg-cyan-950/35 hover:text-cyan-300">Editar</button>
                            <button type="button" data-delete-id="${Number(app.id)}" data-delete-name="${escapedName}" class="rounded-md px-2 py-1 text-[10px] text-slate-700 transition hover:bg-red-950/35 hover:text-red-300">Excluir</button>
                        </div>
                    </div>
                `;

                appsGrid.appendChild(card);
            });

            renderQuickAccess();
            renderCategories();
        }

        function setCategory(category) {
            state.category = category || 'all';
            renderApps();
            closeSidebar();
        }

        async function loadApps() {
            loadingState.classList.remove('hidden');

            try {
                const response = await fetch('/api/apps', { cache: 'no-store' });
                if (!response.ok) throw new Error('Não foi possível carregar os aplicativos.');

                state.apps = await response.json();
                renderApps();

                const editParam = new URLSearchParams(window.location.search).get('edit');
                const requestedEditId = Number(editParam);

                if (
                    editParam &&
                    Number.isInteger(requestedEditId) &&
                    requestedEditId > 0 &&
                    state.apps.some(app => Number(app.id) === requestedEditId)
                ) {
                    openEditModal(requestedEditId);
                    window.history.replaceState({}, '', window.location.pathname);
                }
            } catch (error) {
                showToast(error.message || 'Erro ao carregar aplicativos.', 'error');
            } finally {
                loadingState.classList.add('hidden');
            }
        }

        function resetAppModalState() {
            state.editAppId = null;
            state.currentIcon = '';
            form.reset();
            appIcon.required = false;
            modalTitle.textContent = 'Adicionar aplicativo';
            modalSubtitle.textContent = 'Cadastre uma ferramenta para acesso rápido.';
            saveBtn.textContent = 'Salvar aplicativo';
            iconPreview.innerHTML = '＋';
            iconFileLabel.textContent = 'Selecionar imagem';
            formError.classList.add('hidden');
            formError.textContent = '';
        }

        function openModal() {
            resetAppModalState();
            appId.value = nextSuggestedId();
            appIcon.required = true;

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => appName.focus(), 50);
        }

        function openEditModal(id) {
            const appToEdit = state.apps.find(app => Number(app.id) === Number(id));
            if (!appToEdit) {
                showToast('Aplicativo não encontrado para edição.', 'error');
                return;
            }

            resetAppModalState();
            state.editAppId = Number(appToEdit.id);
            state.currentIcon = String(appToEdit.icon || '');

            modalTitle.textContent = 'Editar aplicativo';
            modalSubtitle.textContent = 'Altere os dados da ferramenta. O ícone só será trocado se você selecionar uma nova imagem.';
            saveBtn.textContent = 'Salvar alterações';

            appId.value = appToEdit.id;
            appName.value = appToEdit.name || '';
            appCategory.value = appToEdit.category || '';
            appUrl.value = appToEdit.url || '';
            appDescription.value = appToEdit.description || '';
            appIcon.required = false;

            iconPreview.innerHTML = iconMarkup(appToEdit.icon, 'h-full w-full object-contain');
            iconFileLabel.textContent = 'Manter ícone atual ou selecionar outro';

            modal.classList.remove('hidden');
            modal.classList.add('flex');
            setTimeout(() => appName.focus(), 50);
        }

        function closeModal() {
            modal.classList.add('hidden');
            modal.classList.remove('flex');
            resetAppModalState();
        }

        function setFormBusy(isBusy) {
            saveBtn.disabled = isBusy;

            if (isBusy) {
                saveBtn.textContent = state.editAppId ? 'Salvando alterações...' : 'Salvando...';
            } else {
                saveBtn.textContent = state.editAppId ? 'Salvar alterações' : 'Salvar aplicativo';
            }
        }

        async function uploadIcon(file) {
            if (!file) throw new Error('Selecione um ícone para o aplicativo.');
            if (file.size > 2 * 1024 * 1024) throw new Error('O ícone deve ter no máximo 2 MB.');

            const response = await fetch('/api/icons', {
                method: 'POST',
                headers: { 'Content-Type': file.type || 'application/octet-stream' },
                body: file
            });

            const data = await response.json();
            if (!response.ok) throw new Error(data.detail || 'Não foi possível enviar o ícone.');
            return data.icon;
        }

        async function saveApp(event) {
            event.preventDefault();

            const editing = state.editAppId !== null;
            const iconFile = appIcon.files?.[0];

            if (!editing && !iconFile) {
                formError.textContent = 'Selecione uma imagem para o ícone do aplicativo.';
                formError.classList.remove('hidden');
                return;
            }

            const currentApp = editing
                ? state.apps.find(app => Number(app.id) === Number(state.editAppId))
                : null;

            const payload = {
                id: editing ? Number(state.editAppId) : Number(appId.value),
                name: appName.value.trim(),
                category: appCategory.value.trim(),
                url: appUrl.value.trim(),
                icon: editing ? state.currentIcon : '',
                description: appDescription.value.trim(),
                favorite: editing ? Boolean(currentApp?.favorite) : false
            };

            formError.classList.add('hidden');
            setFormBusy(true);

            try {
                // Na edição, o ícone atual é mantido se nenhum arquivo novo for escolhido.
                if (iconFile) {
                    payload.icon = await uploadIcon(iconFile);
                }

                const endpoint = editing
                    ? `/api/apps/${state.editAppId}`
                    : '/api/apps';

                const response = await fetch(endpoint, {
                    method: editing ? 'PUT' : 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const data = await response.json();
                if (!response.ok) {
                    const detail = typeof data.detail === 'string'
                        ? data.detail
                        : (editing
                            ? 'Não foi possível atualizar o aplicativo.'
                            : 'Não foi possível cadastrar o aplicativo.');
                    throw new Error(detail);
                }

                if (editing) {
                    const index = state.apps.findIndex(app => Number(app.id) === Number(data.id));
                    if (index !== -1) state.apps[index] = data;
                } else {
                    state.apps.push(data);
                }

                state.category = 'all';
                state.search = '';
                searchInput.value = '';
                renderApps();

                const successMessage = editing
                    ? `“${data.name}” foi atualizado.`
                    : `“${data.name}” foi adicionado ao Hub.`;

                closeModal();
                showToast(successMessage);
            } catch (error) {
                formError.textContent = error.message || 'Erro ao salvar aplicativo.';
                formError.classList.remove('hidden');
            } finally {
                setFormBusy(false);
            }
        }

        async function deleteApp(id, name) {
            const confirmed = window.confirm(`Deseja realmente excluir “${name}” do Hub?`);
            if (!confirmed) return;

            try {
                const response = await fetch(`/api/apps/${id}`, { method: 'DELETE' });
                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Não foi possível excluir o aplicativo.');

                state.apps = state.apps.filter(app => Number(app.id) !== Number(id));

                const validCategories = new Set(state.apps.map(app => app.category));
                if (state.category !== 'all' && !validCategories.has(state.category)) {
                    state.category = 'all';
                }

                renderApps();
                showToast(`“${name}” foi excluído.`);
            } catch (error) {
                showToast(error.message || 'Erro ao excluir aplicativo.', 'error');
            }
        }

        function clearFilters() {
            state.search = '';
            state.category = 'all';
            searchInput.value = '';
            renderApps();
        }

        function openSidebar() {
            sidebar.classList.add('open');
            sidebarOverlay.classList.remove('hidden');
        }

        function closeSidebar() {
            sidebar.classList.remove('open');
            sidebarOverlay.classList.add('hidden');
        }

        profileUploadTrigger.addEventListener('click', () => profileImageInput.click());
        profileImageInput.addEventListener('change', () => {
            const file = profileImageInput.files?.[0];
            if (file) uploadUserProfileImage(file);
        });

        appIcon.addEventListener('change', () => {
            const file = appIcon.files?.[0];

            if (!file) {
                if (state.editAppId && state.currentIcon) {
                    iconPreview.innerHTML = iconMarkup(state.currentIcon, 'h-full w-full object-contain');
                    iconFileLabel.textContent = 'Manter ícone atual ou selecionar outro';
                } else {
                    iconPreview.innerHTML = '＋';
                    iconFileLabel.textContent = 'Selecionar imagem';
                }
                return;
            }

            iconFileLabel.textContent = file.name;

            if (file.type.startsWith('image/')) {
                const reader = new FileReader();
                reader.onload = () => {
                    iconPreview.innerHTML = `<img src="${reader.result}" alt="Prévia do ícone" class="h-full w-full object-contain">`;
                };
                reader.readAsDataURL(file);
            }
        });

        // Eventos gerais
        createExecutableBtn?.addEventListener('click', openBuildExecutableModal);
        showReadmeBtn?.addEventListener('click', openReadmeModal);

        document.getElementById('closeBuildExecutableModalBtn').addEventListener('click', closeBuildExecutableModal);
        document.getElementById('cancelBuildModalBtn').addEventListener('click', closeBuildExecutableModal);
        runExecutableBuildBtn.addEventListener('click', runExecutableBuild);
        openBuildFolderBtn.addEventListener('click', openBuildOutputFolder);
        buildExecutableModal.addEventListener('click', event => {
            if (event.target === buildExecutableModal) closeBuildExecutableModal();
        });

        document.getElementById('closeReadmeModalBtn').addEventListener('click', closeReadmeModal);
        readmeModal.addEventListener('click', event => {
            if (event.target === readmeModal) closeReadmeModal();
        });

        document.getElementById('openModalBtn').addEventListener('click', openModal);
        document.getElementById('sideAddBtn').addEventListener('click', openModal);
        document.getElementById('closeModalX').addEventListener('click', closeModal);
        document.getElementById('cancelBtn').addEventListener('click', closeModal);
        document.getElementById('clearFiltersBtn').addEventListener('click', clearFilters);
        form.addEventListener('submit', saveApp);

        searchInput.addEventListener('input', event => {
            state.search = event.target.value;
            renderApps();
        });

        appsGrid.addEventListener('click', event => {
            const editButton = event.target.closest('[data-edit-id]');
            if (editButton) {
                openEditModal(Number(editButton.dataset.editId));
                return;
            }

            const deleteButton = event.target.closest('[data-delete-id]');
            if (deleteButton) {
                deleteApp(Number(deleteButton.dataset.deleteId), deleteButton.dataset.deleteName);
            }
        });

        sidebarCategories.addEventListener('click', event => {
            const button = event.target.closest('[data-category]');
            if (!button) return;
            setCategory(button.dataset.category);
        });

        categoryChips.addEventListener('click', event => {
            const button = event.target.closest('[data-chip-category]');
            if (!button) return;
            setCategory(button.dataset.chipCategory);
        });

        categoryStats.addEventListener('click', event => {
            const button = event.target.closest('[data-stat-category]');
            if (!button) return;
            setCategory(button.dataset.statCategory);
            document.getElementById('toolsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
        });

        const allToolsButton = document.querySelector('[data-scroll-tools]');
        if (allToolsButton) {
            allToolsButton.addEventListener('click', () => {
                clearFilters();
                document.getElementById('toolsSection').scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
                closeSidebar();
            });
        }

        modal.addEventListener('click', event => {
            if (event.target === modal) closeModal();
        });

        document.addEventListener('keydown', event => {
            if (event.key === 'Escape') {
                if (!modal.classList.contains('hidden')) closeModal();
                closeBuildExecutableModal();
                closeReadmeModal();
                closeSidebar();
            }
        });

        mobileMenuBtn.addEventListener('click', openSidebar);
        sidebarOverlay.addEventListener('click', closeSidebar);

        // Inicialização
        updateClock();
        setInterval(updateClock, 60_000);
        loadUserProfile();
        loadApps();
    </script>
</body>
</html>
'''


# ============================================================
# FRONTEND - GALERIA APPS (ícone + nome)
# ============================================================

APPS_HTML = r'''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>APPS • TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: { hubcyan: '#35d5e6' },
                    boxShadow: { cyan: '0 0 24px rgba(53,213,230,.12)' }
                }
            }
        }
    </script>

    <style>
        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }
        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 78% 5%, rgba(44,196,219,.11), transparent 27rem),
                radial-gradient(circle at 12% 85%, rgba(19,96,129,.10), transparent 26rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }
        body::before {
            content: '';
            position: fixed;
            inset: 0;
            pointer-events: none;
            opacity: .11;
            background-image:
                linear-gradient(rgba(76,188,216,.06) 1px, transparent 1px),
                linear-gradient(90deg, rgba(76,188,216,.05) 1px, transparent 1px);
            background-size: 44px 44px;
            mask-image: linear-gradient(to bottom, black, transparent 86%);
        }
        .glass-top {
            background: rgba(7,24,36,.89);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }
        .sidebar-item { transition: color .18s ease, border-color .18s ease, background .18s ease; }
        .sidebar-item:hover, .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }
        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
            box-shadow: 0 0 14px rgba(53,213,230,.65);
        }
        .app-tile {
            min-height: 146px;
            background: linear-gradient(145deg, rgba(12,35,50,.96), rgba(7,24,36,.98));
            border: 1px solid rgba(129,180,204,.14);
            transition: transform .18s ease, border-color .18s ease, box-shadow .18s ease, background .18s ease;
        }
        .app-tile:hover {
            transform: translateY(-4px);
            border-color: rgba(53,213,230,.42);
            background: linear-gradient(145deg, rgba(14,43,59,.98), rgba(8,29,42,.98));
            box-shadow: 0 15px 36px rgba(0,0,0,.24), 0 0 22px rgba(53,213,230,.055);
        }
        .app-tile[draggable="true"] { cursor: grab; }
        .app-tile[draggable="true"]:active { cursor: grabbing; }
        .app-tile.dragging {
            opacity: .38;
            transform: scale(.96);
            border-color: rgba(53,213,230,.60);
            box-shadow: 0 0 28px rgba(53,213,230,.13);
        }
        .app-tile.drag-target {
            border-color: rgba(53,213,230,.75);
            box-shadow: 0 0 0 2px rgba(53,213,230,.13), 0 0 24px rgba(53,213,230,.10);
        }
        .drag-handle {
            position: absolute;
            right: 11px;
            bottom: 9px;
            color: rgba(148,163,184,.38);
            font-size: 13px;
            line-height: 1;
            letter-spacing: -2px;
            user-select: none;
            pointer-events: none;
        }
        .app-tile:hover .drag-handle { color: rgba(103,232,249,.70); }

        .edit-toggle {
            position: absolute;
            right: 9px;
            top: 8px;
            z-index: 6;
            display: grid;
            width: 30px;
            height: 30px;
            place-items: center;
            border-radius: 9px;
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(5,18,29,.78);
            color: rgba(148,163,184,.55);
            font-size: 16px;
            line-height: 1;
            text-decoration: none;
            transition: color .16s ease, border-color .16s ease, background .16s ease, transform .16s ease;
        }

        .edit-toggle:hover {
            color: #67e8f9;
            border-color: rgba(103,232,249,.34);
            background: rgba(8,145,178,.12);
            transform: scale(1.06);
        }

        .favorite-toggle {
            position: absolute;
            left: 9px;
            top: 8px;
            z-index: 5;
            display: grid;
            width: 30px;
            height: 30px;
            place-items: center;
            border-radius: 9px;
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(5,18,29,.78);
            color: rgba(148,163,184,.48);
            font-size: 17px;
            line-height: 1;
            transition: color .16s ease, border-color .16s ease, background .16s ease, transform .16s ease;
        }

        .favorite-toggle:hover {
            color: #fde68a;
            border-color: rgba(250,204,21,.32);
            background: rgba(113,63,18,.22);
            transform: scale(1.06);
        }

        .favorite-toggle.active {
            color: #facc15;
            border-color: rgba(250,204,21,.34);
            background: rgba(113,63,18,.24);
            text-shadow: 0 0 12px rgba(250,204,21,.30);
        }
        .app-icon {
            width: 64px;
            height: 64px;
            background: linear-gradient(145deg, #071722, #0b2637);
            border: 1px solid rgba(129,180,204,.17);
            box-shadow: inset 0 1px rgba(255,255,255,.025);
            transition: width .18s ease, height .18s ease;
        }

        #appsOnlyGrid.compact-view {
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: .55rem;
        }

        #appsOnlyGrid.compact-view .app-tile {
            min-height: 116px;
            padding: .7rem;
        }

        #appsOnlyGrid.compact-view .app-icon {
            width: 52px;
            height: 52px;
            border-radius: 14px;
        }

        #appsOnlyGrid.compact-view .app-icon img {
            width: 40px !important;
            height: 40px !important;
        }

        #appsOnlyGrid.compact-view .app-name {
            margin-top: .45rem;
            font-size: .76rem;
        }

        #appsOnlyGrid.compact-view .favorite-toggle,
        #appsOnlyGrid.compact-view .edit-toggle {
            width: 26px;
            height: 26px;
            font-size: 14px;
            border-radius: 8px;
        }

        #appsOnlyGrid.compact-view .favorite-toggle {
            left: 7px;
            top: 7px;
        }

        #appsOnlyGrid.compact-view .edit-toggle {
            right: 7px;
            top: 7px;
        }

        #appsOnlyGrid.compact-view .drag-handle {
            right: 8px;
            bottom: 7px;
            font-size: 11px;
        }

        @media (min-width: 640px) {
            #appsOnlyGrid.compact-view {
                grid-template-columns: repeat(4, minmax(0, 1fr));
            }
        }

        @media (min-width: 768px) {
            #appsOnlyGrid.compact-view {
                grid-template-columns: repeat(5, minmax(0, 1fr));
            }
        }

        @media (min-width: 1280px) {
            #appsOnlyGrid.compact-view {
                grid-template-columns: repeat(6, minmax(0, 1fr));
            }
        }

        @media (min-width: 1536px) {
            #appsOnlyGrid.compact-view {
                grid-template-columns: repeat(7, minmax(0, 1fr));
            }
        }
        @media (max-width: 1023px) {
            #appsSidebar { transform: translateX(-100%); transition: transform .22s ease; }
            #appsSidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased selection:bg-cyan-400/20 selection:text-cyan-100">
    <div id="appsSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="appsSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain" onerror="this.style.display='none'; this.nextElementSibling.classList.remove('hidden');">
            <div class="hidden items-center justify-center rounded-2xl border border-cyan-400/20 bg-cyan-400/10 p-4 text-3xl">⬡</div>
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>
            <a href="/apps" class="sidebar-item active relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">▦</span><span class="flex-1 font-semibold">APPS + USADOS</span><span id="appsSidebarCount" class="text-[10px] text-cyan-400/65">0</span>
            </a>
            <a href="/local-apps" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▣</span><span class="flex-1 font-semibold">APPS LOCAIS</span>
            </a>
            <a href="/links" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">★</span><span class="flex-1 font-semibold">LINKS FAVORITOS</span>
            </a>
            <a href="/workspace" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">◈</span><span class="flex-1 font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>
            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span><span class="flex-1 font-semibold">DIAGNÓSTICO WIN</span>
            </a>
            <a href="/" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="flex-1 font-semibold">All tools</span>
            </a>
        </nav>

        <div class="border-t border-slate-700/20 p-3">
            <a href="/docs" target="_blank" rel="noopener noreferrer" class="sidebar-item relative flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⚡</span><span>Documentação API</span>
            </a>
        </div>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex h-[72px] items-center gap-3 px-4 sm:px-6 xl:px-8">
                <button id="appsMobileMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden" aria-label="Abrir menu">☰</button>
                <div class="min-w-0 flex-1">
                    <h1 class="truncate text-base font-bold tracking-[.12em] text-slate-100">APPS</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Galeria de aplicativos</p>
                </div>
                <div class="flex w-full max-w-2xl flex-col gap-2 sm:flex-row sm:items-center">
                    <div class="relative min-w-0 flex-1">
                        <span class="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-600">⌕</span>
                        <input id="appsSearch" type="search" autocomplete="off" placeholder="Buscar app..." class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#0b2232]/85 pl-10 pr-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-600 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                    </div>

                    <div class="relative shrink-0 sm:w-[220px]">
                        <select id="appsCategoryFilter" class="h-10 w-full appearance-none rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-3 pr-9 text-sm text-slate-300 outline-none transition focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                            <option value="all">Todas as categorias</option>
                        </select>
                        <span class="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-[10px] text-slate-500">▼</span>
                    </div>

                    <button
                        id="appsViewToggle"
                        type="button"
                        class="h-10 shrink-0 rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-3 text-xs font-semibold text-slate-400 transition hover:border-cyan-400/25 hover:text-cyan-200"
                        title="Alternar densidade da galeria"
                    >
                        ▦ Compacto
                    </button>
                </div>
            </div>
        </header>

        <main class="mx-auto max-w-[1550px] px-4 py-6 sm:px-6 xl:px-8">
            <div class="mb-5 flex items-end justify-between gap-4">
                <div>
                    <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Aplicativos</p>
                    <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Todos os Apps</h2>
                </div>
                <div class="text-right">
                    <span id="appsCountLabel" class="block text-xs text-slate-600">Carregando...</span>
                    <span id="appsOrderHint" class="mt-1 block text-[10px] text-cyan-400/55">Arraste os cards para organizar</span>
                </div>
            </div>

            <div id="appsLoading" class="py-16 text-center text-sm text-slate-600">Carregando aplicativos...</div>
            <section id="appsOnlyGrid" class="hidden grid grid-cols-2 gap-3 sm:grid-cols-3 md:grid-cols-4 xl:grid-cols-5 2xl:grid-cols-6"></section>
            <div id="appsEmpty" class="hidden rounded-2xl border border-dashed border-slate-700/30 bg-[#081a27]/60 px-6 py-16 text-center">
                <div class="text-4xl">▦</div>
                <p class="mt-3 text-sm font-semibold text-slate-300">Nenhum aplicativo encontrado.</p>
            </div>
        </main>
    </div>

    <div id="appsOrderToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border border-cyan-900/60 bg-[#092633]/95 px-4 py-3 text-sm text-cyan-100 shadow-2xl"></div>

    <script>
        const appsOnlyGrid = document.getElementById('appsOnlyGrid');
        const appsLoading = document.getElementById('appsLoading');
        const appsEmpty = document.getElementById('appsEmpty');
        const appsSearch = document.getElementById('appsSearch');
        const appsCategoryFilter = document.getElementById('appsCategoryFilter');
        const appsCountLabel = document.getElementById('appsCountLabel');
        const appsSidebarCount = document.getElementById('appsSidebarCount');
        const appsSidebar = document.getElementById('appsSidebar');
        const appsSidebarOverlay = document.getElementById('appsSidebarOverlay');
        const appsMobileMenuBtn = document.getElementById('appsMobileMenuBtn');
        const appsOrderHint = document.getElementById('appsOrderHint');
        const appsOrderToast = document.getElementById('appsOrderToast');
        const appsViewToggle = document.getElementById('appsViewToggle');
        let allApps = [];
        let draggedAppId = null;
        let draggedFavorite = null;
        let dragMoved = false;
        let savingOrder = false;

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function iconMarkup(icon) {
            const value = String(icon ?? '').trim();
            const isImage = value.startsWith('/icons/') || /^https?:\/\//i.test(value);
            if (isImage) {
                return `<img src="${escapeHtml(value)}" alt="" class="h-[48px] w-[48px] object-contain" loading="lazy" referrerpolicy="no-referrer" onerror="this.style.display='none';this.parentElement.textContent='🔗';">`;
            }
            return escapeHtml(value || '🔗');
        }

        function safeUrl(value) {
            try {
                const url = new URL(value);
                return ['http:', 'https:'].includes(url.protocol) ? url.href : '#';
            } catch { return '#'; }
        }

        function normalize(value) {
            return String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase();
        }

        function isReorderEnabled() {
            return !appsSearch.value.trim() && (appsCategoryFilter.value || 'all') === 'all';
        }

        function applyAppsViewMode(mode) {
            const compact = mode !== 'comfortable';
            appsOnlyGrid.classList.toggle('compact-view', compact);
            appsViewToggle.textContent = compact ? '▦ Compacto' : '▦ Confortável';
            appsViewToggle.title = compact
                ? 'Clique para usar cards maiores'
                : 'Clique para usar uma galeria mais compacta';

            try {
                localStorage.setItem('techToolHubAppsView', compact ? 'compact' : 'comfortable');
            } catch {}
        }

        function loadAppsViewMode() {
            let saved = 'compact';
            try {
                saved = localStorage.getItem('techToolHubAppsView') || 'compact';
            } catch {}
            applyAppsViewMode(saved);
        }

        function showOrderToast(message, error = false) {
            appsOrderToast.textContent = message;
            appsOrderToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            appsOrderToast.classList.remove('hidden');
            clearTimeout(showOrderToast.timer);
            showOrderToast.timer = setTimeout(() => appsOrderToast.classList.add('hidden'), 2200);
        }

        async function saveAppsOrder() {
            if (savingOrder) return;

            const ids = [...appsOnlyGrid.querySelectorAll('[data-app-id]')]
                .map(tile => Number(tile.dataset.appId));

            if (!ids.length || ids.length !== allApps.length) return;

            savingOrder = true;
            appsOrderHint.textContent = 'Salvando nova ordem...';

            try {
                const response = await fetch('/api/apps/order', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ ids })
                });

                const data = await response.json();
                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível salvar a ordem.');
                }

                const byId = new Map(allApps.map(app => [Number(app.id), app]));
                allApps = ids.map(id => byId.get(id)).filter(Boolean);
                appsOrderHint.textContent = 'Ordem salva • arraste para reorganizar';
                showOrderToast('Ordem dos aplicativos salva.');
            } catch (error) {
                appsOrderHint.textContent = 'Falha ao salvar a ordem';
                showOrderToast(error.message || 'Erro ao salvar a ordem.', true);
                await loadAppsOnly();
            } finally {
                savingOrder = false;
            }
        }

        function clearDragClasses() {
            appsOnlyGrid.querySelectorAll('.app-tile').forEach(tile => {
                tile.classList.remove('dragging', 'drag-target');
            });
        }

        async function toggleFavoriteApp(id) {
            const index = allApps.findIndex(app => Number(app.id) === Number(id));
            if (index === -1) {
                showOrderToast('Aplicativo não encontrado.', true);
                return;
            }

            try {
                const response = await fetch(`/api/apps/${id}/favorite`, {
                    method: 'PATCH'
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível alterar o favorito.');
                }

                allApps[index] = data;
                renderAppsOnly();

                showOrderToast(
                    data.favorite
                        ? `★ ${data.name} adicionado aos favoritos.`
                        : `☆ ${data.name} removido dos favoritos.`
                );
            } catch (error) {
                showOrderToast(error.message || 'Erro ao alterar favorito.', true);
            }
        }

        function populateCategoryFilter() {
            const currentValue = appsCategoryFilter.value || 'all';
            const categories = [...new Set(
                allApps
                    .map(app => String(app.category || 'Outros').trim() || 'Outros')
            )].sort((a, b) => a.localeCompare(b, 'pt-BR'));

            appsCategoryFilter.innerHTML = `
                <option value="all">Todas as categorias</option>
                ${categories.map(category => `
                    <option value="${escapeHtml(category)}">${escapeHtml(category)}</option>
                `).join('')}
            `;

            const stillExists = currentValue === 'all' || categories.includes(currentValue);
            appsCategoryFilter.value = stillExists ? currentValue : 'all';
        }

        function renderAppsOnly() {
            const query = normalize(appsSearch.value.trim());
            const selectedCategory = appsCategoryFilter.value || 'all';
            const normalizedCategory = normalize(selectedCategory);

            const visible = [...allApps]
                .filter(app => {
                    const matchesSearch = !query || normalize(app.name).includes(query);
                    const matchesCategory = selectedCategory === 'all' || normalize(app.category || 'Outros') === normalizedCategory;
                    return matchesSearch && matchesCategory;
                })
                // Favoritos sempre aparecem primeiro. O sort é estável, portanto
                // a ordem manual é preservada dentro de cada grupo.
                .sort((a, b) => Number(Boolean(b.favorite)) - Number(Boolean(a.favorite)));

            const reorderEnabled = isReorderEnabled();

            appsOnlyGrid.innerHTML = visible.map(app => {
                const isFavorite = Boolean(app.favorite);
                const safeName = escapeHtml(app.name);
                const safeAppUrl = escapeHtml(safeUrl(app.url));

                return `
                    <article
                        title="${reorderEnabled ? 'Arraste para organizar ou clique para abrir' : 'Abrir ' + safeName}"
                        data-app-id="${Number(app.id)}"
                        data-favorite="${isFavorite ? '1' : '0'}"
                        draggable="${reorderEnabled ? 'true' : 'false'}"
                        class="app-tile group relative flex flex-col items-center justify-center rounded-xl p-4 text-center"
                    >
                        <button
                            type="button"
                            draggable="false"
                            data-favorite-id="${Number(app.id)}"
                            class="favorite-toggle ${isFavorite ? 'active' : ''}"
                            title="${isFavorite ? 'Remover dos favoritos' : 'Adicionar aos favoritos'}"
                            aria-label="${isFavorite ? 'Remover ' + safeName + ' dos favoritos' : 'Adicionar ' + safeName + ' aos favoritos'}"
                        >${isFavorite ? '★' : '☆'}</button>

                        <a
                            href="/?edit=${Number(app.id)}"
                            draggable="false"
                            class="edit-toggle"
                            title="Editar ${safeName}"
                            aria-label="Editar ${safeName}"
                        >✎</a>

                        ${reorderEnabled ? '<span class="drag-handle" aria-hidden="true">⠿</span>' : ''}

                        <a
                            href="${safeAppUrl}"
                            target="_blank"
                            rel="noopener noreferrer"
                            draggable="false"
                            class="flex w-full flex-1 flex-col items-center justify-center"
                        >
                            <div class="app-icon grid shrink-0 place-items-center overflow-hidden rounded-2xl text-3xl transition group-hover:border-cyan-400/30">${iconMarkup(app.icon)}</div>
                            <div class="app-name mt-2.5 w-full truncate text-sm font-semibold text-slate-300 transition group-hover:text-cyan-100">${safeName}</div>
                        </a>
                    </article>
                `;
            }).join('');

            const favoriteCount = visible.filter(app => Boolean(app.favorite)).length;

            appsOrderHint.textContent = reorderEnabled
                ? `${favoriteCount} favorito${favoriteCount === 1 ? '' : 's'} no topo • arraste dentro de cada grupo para organizar`
                : 'Favoritos ficam no topo • para reorganizar, limpe a busca e selecione “Todas as categorias”';

            const total = allApps.length;
            const hasFilter = query || selectedCategory !== 'all';
            appsSidebarCount.textContent = total;
            appsCountLabel.textContent = hasFilter
                ? `${visible.length} de ${total} apps`
                : `${total} ${total === 1 ? 'app' : 'apps'}`;
            appsOnlyGrid.classList.toggle('hidden', visible.length === 0);
            appsEmpty.classList.toggle('hidden', visible.length !== 0);
        }

        async function loadAppsOnly() {
            try {
                const response = await fetch('/api/apps', { cache: 'no-store' });
                if (!response.ok) throw new Error('Falha ao carregar aplicativos.');
                allApps = await response.json();
                populateCategoryFilter();
                renderAppsOnly();
            } catch (error) {
                appsEmpty.classList.remove('hidden');
                appsEmpty.querySelector('p').textContent = error.message || 'Erro ao carregar aplicativos.';
            } finally {
                appsLoading.classList.add('hidden');
            }
        }

        function closeAppsSidebar() {
            appsSidebar.classList.remove('open');
            appsSidebarOverlay.classList.add('hidden');
        }

        appsSearch.addEventListener('input', renderAppsOnly);
        appsCategoryFilter.addEventListener('change', renderAppsOnly);
        appsViewToggle.addEventListener('click', () => {
            const compact = appsOnlyGrid.classList.contains('compact-view');
            applyAppsViewMode(compact ? 'comfortable' : 'compact');
        });

        appsOnlyGrid.addEventListener('dragstart', event => {
            if (event.target.closest('.edit-toggle, [data-favorite-id], a')) {
                event.preventDefault();
                return;
            }

            const tile = event.target.closest('[data-app-id]');
            if (!tile || !isReorderEnabled()) {
                event.preventDefault();
                return;
            }

            draggedAppId = Number(tile.dataset.appId);
            draggedFavorite = tile.dataset.favorite || '0';
            dragMoved = false;
            tile.classList.add('dragging');

            if (event.dataTransfer) {
                event.dataTransfer.effectAllowed = 'move';
                event.dataTransfer.setData('text/plain', String(draggedAppId));
            }
        });

        appsOnlyGrid.addEventListener('dragover', event => {
            if (draggedAppId === null || !isReorderEnabled()) return;

            event.preventDefault();
            if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';

            const target = event.target.closest('[data-app-id]');
            if (!target || Number(target.dataset.appId) === draggedAppId) return;

            // Favoritos permanecem agrupados no topo. O arraste organiza
            // somente dentro do mesmo grupo (favoritos ou demais apps).
            if ((target.dataset.favorite || '0') !== draggedFavorite) return;

            clearDragClasses();
            const dragged = appsOnlyGrid.querySelector(`[data-app-id="${draggedAppId}"]`);
            if (dragged) dragged.classList.add('dragging');
            target.classList.add('drag-target');

            const rect = target.getBoundingClientRect();
            const before = event.clientY < rect.top + rect.height / 2 ||
                (
                    Math.abs(event.clientY - (rect.top + rect.height / 2)) < rect.height * .22 &&
                    event.clientX < rect.left + rect.width / 2
                );

            if (dragged) {
                appsOnlyGrid.insertBefore(dragged, before ? target : target.nextSibling);
                dragMoved = true;
            }
        });

        appsOnlyGrid.addEventListener('drop', async event => {
            if (draggedAppId === null || !isReorderEnabled()) return;
            event.preventDefault();
            clearDragClasses();

            if (dragMoved) {
                await saveAppsOrder();
            }

            draggedAppId = null;
            draggedFavorite = null;
        });

        appsOnlyGrid.addEventListener('dragend', () => {
            clearDragClasses();
            draggedAppId = null;
            draggedFavorite = null;
            setTimeout(() => { dragMoved = false; }, 80);
        });

        appsOnlyGrid.addEventListener('click', event => {
            const favoriteButton = event.target.closest('[data-favorite-id]');

            if (favoriteButton) {
                event.preventDefault();
                event.stopPropagation();

                if (!dragMoved) {
                    toggleFavoriteApp(Number(favoriteButton.dataset.favoriteId));
                }

                dragMoved = false;
                return;
            }

            if (dragMoved) {
                event.preventDefault();
                event.stopPropagation();
                dragMoved = false;
            }
        });

        appsMobileMenuBtn.addEventListener('click', () => {
            appsSidebar.classList.add('open');
            appsSidebarOverlay.classList.remove('hidden');
        });
        appsSidebarOverlay.addEventListener('click', closeAppsSidebar);
        document.addEventListener('keydown', event => { if (event.key === 'Escape') closeAppsSidebar(); });
        loadAppsViewMode();
        loadAppsOnly();
    </script>
</body>
</html>
'''



# ============================================================
# FRONTEND - LINKS FAVORITOS
# ============================================================

LINKS_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>Favoritos • TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        :root {
            --bg: #06121d;
            --panel: #0a1d2b;
            --line: rgba(129, 180, 204, .14);
            --cyan: #35d5e6;
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }

        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 76% 5%, rgba(44,196,219,.09), transparent 27rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }

        .glass-top {
            background: rgba(7,24,36,.90);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .sidebar-item {
            transition: border-color .18s ease, background .18s ease, color .18s ease;
        }

        .sidebar-item:hover,
        .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
        }

        .category-section {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(145deg, rgba(13,35,50,.93), rgba(8,25,37,.96));
        }

        .favorite-row {
            transition: transform .16s ease, border-color .16s ease, background .16s ease, opacity .16s ease;
        }

        .favorite-row:hover {
            border-color: rgba(53,213,230,.26);
            background: rgba(15,45,61,.78);
        }

        .favorite-row[draggable="true"] { cursor: grab; }
        .favorite-row[draggable="true"]:active { cursor: grabbing; }

        .favorite-row.dragging {
            opacity: .34;
            transform: scale(.985);
            border-color: rgba(53,213,230,.55);
        }

        .favorite-row.drag-target {
            border-color: rgba(53,213,230,.75);
            box-shadow: 0 0 0 2px rgba(53,213,230,.10);
        }

        .drag-handle {
            color: rgba(148,163,184,.38);
            letter-spacing: -2px;
            user-select: none;
        }

        .favorite-row:hover .drag-handle { color: rgba(103,232,249,.75); }

        .favicon {
            background: #071722;
            border: 1px solid rgba(129,180,204,.16);
        }

        #linkDropZone.drag-over {
            border-color: rgba(53,213,230,.70);
            background: rgba(53,213,230,.09);
            box-shadow: 0 0 0 2px rgba(53,213,230,.07), 0 0 30px rgba(53,213,230,.08);
        }

        .modal-backdrop {
            background: rgba(1,8,13,.80);
            backdrop-filter: blur(7px);
            -webkit-backdrop-filter: blur(7px);
        }

        @media (max-width: 1023px) {
            #linksSidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }
            #linksSidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased selection:bg-cyan-400/20 selection:text-cyan-100">
    <div id="linksSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="linksSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain" onerror="this.style.display='none'; this.nextElementSibling.classList.remove('hidden');">
            <div class="hidden items-center justify-center rounded-2xl border border-cyan-400/20 bg-cyan-400/10 p-4 text-3xl">⬡</div>
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▦</span><span class="font-semibold">APPS + USADOS</span>
            </a>

            <a href="/local-apps" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▣</span><span class="font-semibold">APPS LOCAIS</span>
            </a>

            <a href="/links" class="sidebar-item active relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">★</span>
                <span class="flex-1 font-semibold">LINKS FAVORITOS</span>
                <span id="linksSidebarCount" class="text-[10px] text-cyan-400/65">0</span>
            </a>

            <a href="/workspace" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">◈</span><span class="font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>

            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span><span class="font-semibold">DIAGNÓSTICO WIN</span>
            </a>

            <a href="/" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="font-semibold">All tools</span>
            </a>
        </nav>

        <div class="border-t border-slate-700/20 p-3">
            <a href="/docs" target="_blank" rel="noopener noreferrer" class="sidebar-item relative flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⚡</span><span>Documentação API</span>
            </a>
        </div>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex min-h-[72px] flex-wrap items-center gap-3 px-4 py-3 sm:px-6 xl:px-8">
                <button id="linksMobileMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden" aria-label="Abrir menu">☰</button>

                <div class="min-w-[180px] flex-1">
                    <h1 class="truncate text-base font-bold tracking-[.12em] text-slate-100">FAVORITOS</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Links organizados por categoria</p>
                </div>

                <div class="flex w-full flex-col gap-2 md:w-auto md:flex-row md:items-center">
                    <div class="relative md:w-[290px]">
                        <span class="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-600">⌕</span>
                        <input id="linksSearch" type="search" autocomplete="off" placeholder="Buscar link..." class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#0b2232]/85 pl-10 pr-3 text-sm text-slate-200 outline-none transition placeholder:text-slate-600 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                    </div>

                    <div class="relative md:w-[220px]">
                        <select id="linksCategoryFilter" class="h-10 w-full appearance-none rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-3 pr-9 text-sm text-slate-300 outline-none transition focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                            <option value="all">Todas as categorias</option>
                        </select>
                        <span class="pointer-events-none absolute right-3 top-1/2 -translate-y-1/2 text-[10px] text-slate-500">▼</span>
                    </div>

                    <button id="importBrowserLinksBtn" type="button" disabled class="hidden">
                        Importação desativada
                    </button>

                    <button id="openLinkModalBtn" type="button" class="h-10 rounded-xl border border-cyan-300/20 bg-cyan-400/10 px-4 text-sm font-bold text-cyan-200 transition hover:border-cyan-300/40 hover:bg-cyan-400/15">
                        ＋ Adicionar Link
                    </button>
                </div>
            </div>
        </header>

        <main class="mx-auto max-w-[1350px] px-4 py-6 sm:px-6 xl:px-8">
            <section class="mb-5 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
                <div>
                    <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Sessão temporária</p>
                    <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Links temporários</h2>
                    <p class="mt-1 max-w-2xl text-sm text-slate-500">
                        Os links adicionados aqui ficam somente na memória enquanto o TECH TOOL HUB estiver aberto. Nada é salvo em disco.
                    </p>
                </div>

                <div class="text-right">
                    <div id="linksCountLabel" class="text-xs text-slate-600">Carregando...</div>
                    <div id="linksOrderHint" class="mt-1 text-[10px] text-cyan-400/55">Arraste os links para organizar</div>
                </div>
            </section>

            <section class="mb-5 rounded-2xl border border-emerald-400/20 bg-emerald-400/[.035] px-5 py-4">
                <div class="text-sm font-bold text-emerald-200">🔒 Política zero-credenciais</div>
                <div class="mt-1 text-xs leading-5 text-slate-500">
                    O Hub não armazena senhas, tokens, cookies, sessões ou links pessoais. Ao abrir um site, a autenticação depende exclusivamente da sessão existente no navegador/site.
                </div>
            </section>

            <section id="linkDropZone" class="mb-5 rounded-2xl border border-dashed border-cyan-400/20 bg-cyan-400/[.035] px-5 py-5 transition">
                <div class="flex flex-col items-center justify-center gap-2 text-center sm:flex-row sm:text-left">
                    <div class="grid h-11 w-11 shrink-0 place-items-center rounded-xl border border-cyan-400/15 bg-cyan-400/5 text-xl text-cyan-300">↧</div>
                    <div class="min-w-0">
                        <div class="text-sm font-semibold text-slate-300">Arraste um link aqui para usar nesta sessão</div>
                        <div class="mt-0.5 text-xs text-slate-600">O atalho fica somente na memória e desaparece quando o Hub for encerrado.</div>
                    </div>
                </div>
            </section>

            <div id="linksLoading" class="py-16 text-center text-sm text-slate-600">Carregando favoritos...</div>

            <div id="linksGroups" class="hidden space-y-5"></div>

            <div id="linksEmpty" class="hidden rounded-2xl border border-dashed border-slate-700/30 bg-[#081a27]/60 px-6 py-16 text-center">
                <div class="text-4xl">★</div>
                <p class="mt-3 text-sm font-semibold text-slate-300">Nenhum link encontrado.</p>
                <p class="mt-1 text-xs text-slate-600">Use “Adicionar Link” para criar um atalho temporário nesta sessão.</p>
            </div>
        </main>
    </div>

    <!-- MODAL IMPORTAR FAVORITOS DO NAVEGADOR -->
    <div id="browserImportModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-xl rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 class="text-xl font-bold text-slate-100">Importação desativada</h3>
                    <p class="mt-1 text-xs text-slate-500">A política de segurança impede leitura e armazenamento de favoritos pessoais do navegador.</p>
                </div>
                <button id="closeBrowserImportModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <div id="browserImportLoading" class="rounded-xl border border-slate-700/20 bg-[#061722] px-4 py-8 text-center text-xs text-slate-600">
                Procurando favoritos nos navegadores...
            </div>

            <div id="browserImportList" class="hidden space-y-2"></div>

            <label class="mt-4 flex cursor-pointer items-start gap-3 rounded-xl border border-slate-700/20 bg-[#061722] p-3">
                <input id="browserUseFolders" type="checkbox" checked class="mt-0.5 h-4 w-4 accent-cyan-400">
                <span>
                    <span class="block text-xs font-semibold text-slate-300">Usar pastas como categorias</span>
                    <span class="mt-0.5 block text-[10px] leading-4 text-slate-600">Ex.: uma pasta “Cursos” do navegador vira a categoria “Cursos” no Hub.</span>
                </span>
            </label>

            <div id="browserImportError" class="mt-4 hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

            <div class="mt-5 flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
                <button id="cancelBrowserImportBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 hover:bg-slate-800/35 hover:text-slate-200">
                    Cancelar
                </button>
                <button id="runBrowserImportBtn" type="button" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15 disabled:cursor-not-allowed disabled:opacity-50">
                    Importar selecionados
                </button>
            </div>
        </div>
    </div>

    <!-- MODAL ADICIONAR LINK -->
    <div id="linkModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-xl rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 id="linkModalTitle" class="text-xl font-bold text-slate-100">Adicionar link favorito</h3>
                    <p id="linkModalSubtitle" class="mt-1 text-xs text-slate-500">Escolha um nome, URL e grupo/categoria.</p>
                </div>
                <button id="closeLinkModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <form id="linkForm" class="space-y-4">
                <div>
                    <label for="linkTitle" class="mb-1.5 block text-xs font-semibold text-slate-400">Nome do link</label>
                    <input id="linkTitle" type="text" maxlength="140" required placeholder="Ex.: Tutorial FastAPI" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                </div>

                <div>
                    <label for="linkUrl" class="mb-1.5 block text-xs font-semibold text-slate-400">URL</label>
                    <input id="linkUrl" type="url" maxlength="1000" required placeholder="https://..." class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                </div>

                <div>
                    <label for="linkCategory" class="mb-1.5 block text-xs font-semibold text-slate-400">Categoria / Grupo</label>
                    <input id="linkCategory" type="text" list="linkCategorySuggestions" maxlength="60" required placeholder="Ex.: Tecnologia" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10">
                    <datalist id="linkCategorySuggestions">
                        <option value="Vídeos">
                        <option value="Imagens">
                        <option value="Filmes">
                        <option value="Livros">
                        <option value="Tecnologia">
                        <option value="Aplicativos">
                        <option value="Dicas">
                        <option value="Interessante">
                        <option value="Estudos">
                        <option value="Notícias">
                        <option value="Trabalho">
                        <option value="Compras">
                    </datalist>
                </div>

                <div>
                    <label for="linkNote" class="mb-1.5 block text-xs font-semibold text-slate-400">Observação <span class="font-normal text-slate-600">(opcional)</span></label>
                    <textarea id="linkNote" rows="3" maxlength="300" placeholder="Uma anotação rápida sobre esse link..." class="w-full resize-none rounded-xl border border-slate-700/30 bg-[#06131d] px-3 py-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35 focus:ring-2 focus:ring-cyan-400/10"></textarea>
                </div>

                <div id="linkFormError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:justify-end">
                    <button id="cancelLinkBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 hover:bg-slate-800/35 hover:text-slate-200">Cancelar</button>
                    <button id="saveLinkBtn" type="submit" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 hover:border-cyan-300/35 hover:bg-cyan-400/15 disabled:cursor-not-allowed disabled:opacity-50">Salvar link</button>
                </div>
            </form>
        </div>
    </div>

    <div id="linksToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const PRESET_CATEGORIES = [
            'Vídeos', 'Imagens', 'Filmes', 'Livros', 'Tecnologia', 'Aplicativos',
            'Dicas', 'Interessante', 'Estudos', 'Notícias', 'Trabalho', 'Compras'
        ];

        const linksGroups = document.getElementById('linksGroups');
        const linksLoading = document.getElementById('linksLoading');
        const linksEmpty = document.getElementById('linksEmpty');
        const linksSearch = document.getElementById('linksSearch');
        const linksCategoryFilter = document.getElementById('linksCategoryFilter');
        const linksCountLabel = document.getElementById('linksCountLabel');
        const linksSidebarCount = document.getElementById('linksSidebarCount');
        const linksOrderHint = document.getElementById('linksOrderHint');
        const linksToast = document.getElementById('linksToast');
        const linkDropZone = document.getElementById('linkDropZone');

        const browserImportModal = document.getElementById('browserImportModal');
        const browserImportLoading = document.getElementById('browserImportLoading');
        const browserImportList = document.getElementById('browserImportList');
        const browserUseFolders = document.getElementById('browserUseFolders');
        const browserImportError = document.getElementById('browserImportError');
        const runBrowserImportBtn = document.getElementById('runBrowserImportBtn');

        const linksSidebar = document.getElementById('linksSidebar');
        const linksSidebarOverlay = document.getElementById('linksSidebarOverlay');
        const linksMobileMenuBtn = document.getElementById('linksMobileMenuBtn');

        const linkModal = document.getElementById('linkModal');
        const linkForm = document.getElementById('linkForm');
        const linkModalTitle = document.getElementById('linkModalTitle');
        const linkModalSubtitle = document.getElementById('linkModalSubtitle');
        const linkTitle = document.getElementById('linkTitle');
        const linkUrl = document.getElementById('linkUrl');
        const linkCategory = document.getElementById('linkCategory');
        const linkNote = document.getElementById('linkNote');
        const linkFormError = document.getElementById('linkFormError');
        const saveLinkBtn = document.getElementById('saveLinkBtn');

        let allLinks = [];
        let draggedLinkId = null;
        let draggedCategory = null;
        let dragMoved = false;
        let savingOrder = false;
        let editLinkId = null;

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function normalize(value) {
            return String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase();
        }

        function safeUrl(value) {
            try {
                const url = new URL(value);
                return ['http:', 'https:'].includes(url.protocol) ? url.href : '#';
            } catch {
                return '#';
            }
        }

        function hostFromUrl(value) {
            try {
                return new URL(value).hostname.replace(/^www\./, '');
            } catch {
                return '';
            }
        }

        function titleFromDroppedHtml(html) {
            if (!html) return '';

            try {
                const doc = new DOMParser().parseFromString(html, 'text/html');
                const anchor = doc.querySelector('a[href]');
                return String(anchor?.textContent || '').trim();
            } catch {
                return '';
            }
        }

        function droppedLinkData(dataTransfer) {
            const uriList = String(dataTransfer.getData('text/uri-list') || '')
                .split(/\r?\n/)
                .map(line => line.trim())
                .find(line => line && !line.startsWith('#'));

            const plain = String(dataTransfer.getData('text/plain') || '').trim();
            const html = String(dataTransfer.getData('text/html') || '');

            const candidates = [uriList, plain];

            for (const candidate of candidates) {
                if (!candidate) continue;

                try {
                    const url = new URL(candidate);
                    if (!['http:', 'https:'].includes(url.protocol)) continue;

                    return {
                        url: url.href,
                        title: titleFromDroppedHtml(html) || url.hostname.replace(/^www\./, '')
                    };
                } catch {
                    // O texto arrastado pode ser apenas o nome do link.
                }
            }

            if (html) {
                try {
                    const doc = new DOMParser().parseFromString(html, 'text/html');
                    const anchor = doc.querySelector('a[href]');
                    const href = anchor?.getAttribute('href');

                    if (href) {
                        const url = new URL(href);
                        if (['http:', 'https:'].includes(url.protocol)) {
                            return {
                                url: url.href,
                                title: String(anchor.textContent || '').trim() || url.hostname.replace(/^www\./, '')
                            };
                        }
                    }
                } catch {}
            }

            return null;
        }

        function faviconUrl(value) {
            const url = safeUrl(value);
            if (url === '#') return '';
            return `https://www.google.com/s2/favicons?domain_url=${encodeURIComponent(url)}&sz=64`;
        }

        function showToast(message, error = false) {
            linksToast.textContent = message;
            linksToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            linksToast.classList.remove('hidden');
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => linksToast.classList.add('hidden'), 2400);
        }

        function isReorderEnabled() {
            return !linksSearch.value.trim() && (linksCategoryFilter.value || 'all') === 'all';
        }

        function getCategories() {
            const dynamic = allLinks
                .map(link => String(link.category || '').trim())
                .filter(Boolean);

            return [...new Set([...PRESET_CATEGORIES, ...dynamic])]
                .sort((a, b) => a.localeCompare(b, 'pt-BR'));
        }

        function populateCategoryFilter() {
            const current = linksCategoryFilter.value || 'all';

            linksCategoryFilter.innerHTML =
                '<option value="all">Todas as categorias</option>' +
                getCategories()
                    .filter(category => allLinks.some(link => link.category === category))
                    .map(category => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`)
                    .join('');

            if ([...linksCategoryFilter.options].some(option => option.value === current)) {
                linksCategoryFilter.value = current;
            }
        }

        function visibleLinks() {
            const query = normalize(linksSearch.value.trim());
            const category = linksCategoryFilter.value || 'all';

            return allLinks.filter(link => {
                const categoryMatch = category === 'all' || link.category === category;
                if (!categoryMatch) return false;

                if (!query) return true;

                return [
                    link.title,
                    link.category,
                    link.note,
                    hostFromUrl(link.url),
                    link.url
                ].some(value => normalize(value).includes(query));
            });
        }

        function groupLinks(links) {
            const map = new Map();

            links.forEach(link => {
                const category = String(link.category || 'Outros').trim() || 'Outros';
                if (!map.has(category)) map.set(category, []);
                map.get(category).push(link);
            });

            return [...map.entries()]
                .sort((a, b) => a[0].localeCompare(b[0], 'pt-BR'));
        }

        function rowMarkup(link, reorderEnabled) {
            const url = safeUrl(link.url);
            const host = hostFromUrl(link.url);
            const icon = faviconUrl(link.url);
            const initial = escapeHtml(String(link.title || '?').trim().charAt(0).toUpperCase() || '?');

            return `
                <article
                    class="favorite-row relative flex items-center gap-3 rounded-xl border border-slate-700/20 bg-[#091b28]/80 p-3"
                    data-link-id="${Number(link.id)}"
                    data-category="${escapeHtml(link.category)}"
                    draggable="${reorderEnabled ? 'true' : 'false'}"
                >
                    <div class="drag-handle w-4 shrink-0 text-center text-sm" title="${reorderEnabled ? 'Arraste para organizar' : ''}">⠿</div>

                    <a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="favicon grid h-11 w-11 shrink-0 place-items-center overflow-hidden rounded-lg text-sm font-bold text-cyan-300">
                        ${icon ? `<img src="${escapeHtml(icon)}" alt="" class="h-7 w-7 object-contain" onerror="this.remove(); this.parentElement.textContent='${initial}'">` : initial}
                    </a>

                    <a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="min-w-0 flex-1">
                        <div class="truncate text-sm font-semibold text-slate-200 hover:text-cyan-200">${escapeHtml(link.title)}</div>
                        <div class="mt-0.5 truncate text-[11px] text-slate-600">${escapeHtml(host || link.url)}</div>
                        ${link.note ? `<div class="mt-1 truncate text-[11px] text-slate-500">${escapeHtml(link.note)}</div>` : ''}
                    </a>

                    <div class="flex shrink-0 items-center gap-1">
                        <a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer" class="rounded-lg px-2.5 py-2 text-[11px] font-semibold text-cyan-400/70 hover:bg-cyan-400/5 hover:text-cyan-200">Abrir ↗</a>
                        <button type="button" data-edit-link="${Number(link.id)}" class="rounded-lg px-2 py-2 text-[11px] text-slate-500 hover:bg-cyan-950/30 hover:text-cyan-300" title="Editar">Editar</button>
                        <button type="button" data-delete-link="${Number(link.id)}" data-delete-title="${escapeHtml(link.title)}" class="rounded-lg px-2 py-2 text-[11px] text-slate-600 hover:bg-red-950/30 hover:text-red-300" title="Excluir">✕</button>
                    </div>
                </article>
            `;
        }

        function renderLinks() {
            const visible = visibleLinks();
            const grouped = groupLinks(visible);
            const reorderEnabled = isReorderEnabled();

            linksSidebarCount.textContent = allLinks.length;
            linksCountLabel.textContent = `${visible.length} de ${allLinks.length} ${allLinks.length === 1 ? 'link' : 'links'}`;

            linksGroups.innerHTML = grouped.map(([category, items]) => `
                <section class="category-section rounded-2xl p-4" data-category-section="${escapeHtml(category)}">
                    <div class="mb-3 flex items-center justify-between gap-3 border-b border-slate-700/20 pb-3">
                        <div class="flex min-w-0 items-center gap-2">
                            <span class="grid h-7 w-7 shrink-0 place-items-center rounded-lg border border-cyan-400/10 bg-cyan-400/5 text-xs text-cyan-300">#</span>
                            <h3 class="truncate text-sm font-bold text-slate-200">${escapeHtml(category)}</h3>
                        </div>
                        <span class="rounded-md border border-slate-700/20 bg-[#061722] px-2 py-1 text-[10px] text-slate-600">${items.length}</span>
                    </div>

                    <div class="links-category-list space-y-2" data-category-list="${escapeHtml(category)}">
                        ${items.map(link => rowMarkup(link, reorderEnabled)).join('')}
                    </div>
                </section>
            `).join('');

            const hasVisible = visible.length > 0;
            linksGroups.classList.toggle('hidden', !hasVisible);
            linksEmpty.classList.toggle('hidden', hasVisible);

            linksOrderHint.textContent = reorderEnabled
                ? 'Arraste links dentro da mesma categoria • a ordem é salva automaticamente'
                : 'Para reorganizar, limpe a busca e selecione “Todas as categorias”';
        }

        async function loadLinks() {
            linksLoading.classList.remove('hidden');
            linksGroups.classList.add('hidden');
            linksEmpty.classList.add('hidden');

            try {
                const response = await fetch('/api/links', { cache: 'no-store' });
                if (!response.ok) throw new Error('Não foi possível carregar os favoritos.');

                allLinks = await response.json();
                populateCategoryFilter();
                renderLinks();
            } catch (error) {
                showToast(error.message || 'Erro ao carregar favoritos.', true);
            } finally {
                linksLoading.classList.add('hidden');
            }
        }

        function closeBrowserImportModal() {
            browserImportModal.classList.add('hidden');
            browserImportModal.classList.remove('flex');
            browserImportError.classList.add('hidden');
            browserImportError.textContent = '';
        }

        async function openBrowserImportModal() {
            browserImportModal.classList.remove('hidden');
            browserImportModal.classList.add('flex');
            browserImportLoading.classList.remove('hidden');
            browserImportList.classList.add('hidden');
            browserImportList.innerHTML = '';
            browserImportError.classList.add('hidden');
            runBrowserImportBtn.disabled = true;

            try {
                const response = await fetch('/api/links/browser-import/status', {
                    cache: 'no-store'
                });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível procurar favoritos.');
                }

                browserImportList.innerHTML = (data.browsers || []).map(browser => `
                    <label class="flex items-center gap-3 rounded-xl border border-slate-700/20 bg-[#061722] p-3 ${browser.count ? 'cursor-pointer' : 'opacity-45'}">
                        <input
                            type="checkbox"
                            class="browser-import-choice h-4 w-4 accent-cyan-400"
                            value="${escapeHtml(browser.id)}"
                            ${browser.count ? 'checked' : 'disabled'}
                        >
                        <span class="grid h-9 w-9 place-items-center rounded-lg border border-slate-700/20 bg-[#0a1d2b] text-xl">${escapeHtml(browser.icon)}</span>
                        <span class="min-w-0 flex-1">
                            <span class="block text-xs font-semibold text-slate-300">${escapeHtml(browser.name)}</span>
                            <span class="mt-0.5 block text-[10px] ${browser.count ? 'text-cyan-400/65' : 'text-slate-700'}">
                                ${browser.count ? `${browser.count} favorito${browser.count === 1 ? '' : 's'} encontrado${browser.count === 1 ? '' : 's'}` : 'Nenhum favorito encontrado'}
                            </span>
                        </span>
                    </label>
                `).join('');

                browserImportLoading.classList.add('hidden');
                browserImportList.classList.remove('hidden');
                updateBrowserImportButton();
            } catch (error) {
                browserImportLoading.classList.add('hidden');
                browserImportError.textContent = error.message || 'Erro ao detectar favoritos.';
                browserImportError.classList.remove('hidden');
            }
        }

        function selectedBrowserImports() {
            return [...document.querySelectorAll('.browser-import-choice:checked')]
                .map(input => input.value);
        }

        function updateBrowserImportButton() {
            runBrowserImportBtn.disabled = selectedBrowserImports().length === 0;
        }

        async function runBrowserImport() {
            const browsers = selectedBrowserImports();

            if (!browsers.length) {
                browserImportError.textContent = 'Selecione ao menos um navegador.';
                browserImportError.classList.remove('hidden');
                return;
            }

            runBrowserImportBtn.disabled = true;
            runBrowserImportBtn.textContent = 'Importando...';
            browserImportError.classList.add('hidden');

            try {
                const response = await fetch('/api/links/browser-import', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        browsers,
                        use_folders: browserUseFolders.checked
                    })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível importar os favoritos.');
                }

                closeBrowserImportModal();
                await loadLinks();

                const parts = [`${data.imported} importado${data.imported === 1 ? '' : 's'}`];
                if (data.duplicates) parts.push(`${data.duplicates} duplicado${data.duplicates === 1 ? '' : 's'} ignorado${data.duplicates === 1 ? '' : 's'}`);
                if (data.invalid) parts.push(`${data.invalid} inválido${data.invalid === 1 ? '' : 's'} ignorado${data.invalid === 1 ? '' : 's'}`);

                showToast(`Importação concluída: ${parts.join(' • ')}.`);
            } catch (error) {
                browserImportError.textContent = error.message || 'Erro durante a importação.';
                browserImportError.classList.remove('hidden');
            } finally {
                runBrowserImportBtn.disabled = false;
                runBrowserImportBtn.textContent = 'Importar selecionados';
            }
        }

        function resetLinkModal() {
            editLinkId = null;
            linkForm.reset();
            linkFormError.classList.add('hidden');
            linkFormError.textContent = '';
            linkModalTitle.textContent = 'Adicionar link favorito';
            linkModalSubtitle.textContent = 'Escolha um nome, URL e grupo/categoria.';
            saveLinkBtn.textContent = 'Salvar link';
        }

        function openLinkModal(prefill = null) {
            resetLinkModal();

            if (prefill) {
                linkTitle.value = prefill.title || '';
                linkUrl.value = prefill.url || '';
            }

            linkModal.classList.remove('hidden');
            linkModal.classList.add('flex');

            setTimeout(() => {
                if (prefill?.url && !prefill?.title) linkTitle.focus();
                else if (prefill?.url) linkCategory.focus();
                else linkTitle.focus();
            }, 50);
        }

        function openEditLinkModal(id) {
            const link = allLinks.find(item => Number(item.id) === Number(id));

            if (!link) {
                showToast('Link não encontrado para edição.', true);
                return;
            }

            resetLinkModal();
            editLinkId = Number(link.id);

            linkModalTitle.textContent = 'Editar link favorito';
            linkModalSubtitle.textContent = 'Altere nome, endereço, categoria ou observação.';
            saveLinkBtn.textContent = 'Salvar alterações';

            linkTitle.value = link.title || '';
            linkUrl.value = link.url || '';
            linkCategory.value = link.category || '';
            linkNote.value = link.note || '';

            linkModal.classList.remove('hidden');
            linkModal.classList.add('flex');
            setTimeout(() => linkTitle.focus(), 50);
        }

        function closeLinkModal() {
            linkModal.classList.add('hidden');
            linkModal.classList.remove('flex');
            resetLinkModal();
        }

        async function saveLink(event) {
            event.preventDefault();

            const editing = editLinkId !== null;

            const payload = {
                title: linkTitle.value.trim(),
                category: linkCategory.value.trim(),
                url: linkUrl.value.trim(),
                note: linkNote.value.trim()
            };

            saveLinkBtn.disabled = true;
            saveLinkBtn.textContent = editing ? 'Salvando alterações...' : 'Salvando...';
            linkFormError.classList.add('hidden');

            try {
                const response = await fetch(
                    editing ? `/api/links/${editLinkId}` : '/api/links',
                    {
                        method: editing ? 'PUT' : 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    }
                );

                const data = await response.json();

                if (!response.ok) {
                    const detail = typeof data.detail === 'string'
                        ? data.detail
                        : (editing
                            ? 'Não foi possível atualizar o link.'
                            : 'Não foi possível salvar o link.');
                    throw new Error(detail);
                }

                if (editing) {
                    const index = allLinks.findIndex(item => Number(item.id) === Number(data.id));
                    if (index !== -1) allLinks[index] = data;
                } else {
                    allLinks.push(data);
                }

                linksSearch.value = '';
                linksCategoryFilter.value = 'all';
                populateCategoryFilter();
                renderLinks();

                const message = editing
                    ? `“${data.title}” foi atualizado.`
                    : `“${data.title}” foi adicionado aos favoritos.`;

                closeLinkModal();
                showToast(message);
            } catch (error) {
                linkFormError.textContent = error.message || 'Erro ao salvar link.';
                linkFormError.classList.remove('hidden');
            } finally {
                saveLinkBtn.disabled = false;
                saveLinkBtn.textContent = editLinkId ? 'Salvar alterações' : 'Salvar link';
            }
        }

        async function deleteLink(id, title) {
            if (!window.confirm(`Deseja realmente excluir “${title}” dos favoritos?`)) return;

            try {
                const response = await fetch(`/api/links/${id}`, { method: 'DELETE' });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível excluir o link.');
                }

                allLinks = allLinks.filter(link => Number(link.id) !== Number(id));
                populateCategoryFilter();
                renderLinks();
                showToast(`“${title}” foi excluído.`);
            } catch (error) {
                showToast(error.message || 'Erro ao excluir link.', true);
            }
        }

        function clearDragState() {
            document.querySelectorAll('.favorite-row').forEach(row => {
                row.classList.remove('dragging', 'drag-target');
            });
        }

        async function saveOrderFromDom() {
            if (savingOrder) return;

            const ids = [...document.querySelectorAll('#linksGroups .favorite-row[data-link-id]')]
                .map(row => Number(row.dataset.linkId));

            if (!ids.length || ids.length !== allLinks.length) return;

            savingOrder = true;
            linksOrderHint.textContent = 'Salvando nova ordem...';

            try {
                const response = await fetch('/api/links/order', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ ids })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível salvar a ordem.');
                }

                const byId = new Map(allLinks.map(link => [Number(link.id), link]));
                allLinks = ids.map(id => byId.get(id)).filter(Boolean);
                linksOrderHint.textContent = 'Ordem salva • arraste para reorganizar';
                showToast('Ordem dos favoritos salva.');
            } catch (error) {
                showToast(error.message || 'Erro ao salvar a ordem.', true);
                await loadLinks();
            } finally {
                savingOrder = false;
            }
        }

        linksGroups.addEventListener('dragstart', event => {
            const row = event.target.closest('.favorite-row[data-link-id]');
            if (!row || !isReorderEnabled()) {
                event.preventDefault();
                return;
            }

            draggedLinkId = Number(row.dataset.linkId);
            draggedCategory = row.dataset.category;
            dragMoved = false;
            row.classList.add('dragging');

            if (event.dataTransfer) {
                event.dataTransfer.effectAllowed = 'move';
                event.dataTransfer.setData('text/plain', String(draggedLinkId));
            }
        });

        linksGroups.addEventListener('dragover', event => {
            if (draggedLinkId === null || !isReorderEnabled()) return;

            const target = event.target.closest('.favorite-row[data-link-id]');
            if (!target) return;
            if (Number(target.dataset.linkId) === draggedLinkId) return;

            // A organização é feita dentro do próprio grupo/categoria.
            if (target.dataset.category !== draggedCategory) return;

            event.preventDefault();
            if (event.dataTransfer) event.dataTransfer.dropEffect = 'move';

            clearDragState();

            const dragged = document.querySelector(`.favorite-row[data-link-id="${draggedLinkId}"]`);
            if (dragged) dragged.classList.add('dragging');
            target.classList.add('drag-target');

            const rect = target.getBoundingClientRect();
            const before = event.clientY < rect.top + rect.height / 2;

            if (dragged && target.parentElement === dragged.parentElement) {
                target.parentElement.insertBefore(dragged, before ? target : target.nextSibling);
                dragMoved = true;
            }
        });

        linksGroups.addEventListener('drop', async event => {
            if (draggedLinkId === null || !isReorderEnabled()) return;

            const target = event.target.closest('.favorite-row[data-link-id]');
            if (!target || target.dataset.category !== draggedCategory) return;

            event.preventDefault();
            clearDragState();

            if (dragMoved) await saveOrderFromDom();

            draggedLinkId = null;
            draggedCategory = null;
        });

        linksGroups.addEventListener('dragend', () => {
            clearDragState();
            draggedLinkId = null;
            draggedCategory = null;
            setTimeout(() => { dragMoved = false; }, 80);
        });

        linksGroups.addEventListener('click', event => {
            const editButton = event.target.closest('[data-edit-link]');
            if (editButton) {
                event.preventDefault();
                event.stopPropagation();
                openEditLinkModal(Number(editButton.dataset.editLink));
                return;
            }

            const deleteButton = event.target.closest('[data-delete-link]');
            if (deleteButton) {
                event.preventDefault();
                event.stopPropagation();
                deleteLink(Number(deleteButton.dataset.deleteLink), deleteButton.dataset.deleteTitle);
                return;
            }

            if (dragMoved) {
                event.preventDefault();
                event.stopPropagation();
                dragMoved = false;
            }
        });

        ['dragenter', 'dragover'].forEach(eventName => {
            linkDropZone.addEventListener(eventName, event => {
                event.preventDefault();
                event.stopPropagation();

                if (event.dataTransfer) {
                    event.dataTransfer.dropEffect = 'copy';
                }

                linkDropZone.classList.add('drag-over');
            });
        });

        ['dragleave', 'drop'].forEach(eventName => {
            linkDropZone.addEventListener(eventName, event => {
                event.preventDefault();
                event.stopPropagation();
                linkDropZone.classList.remove('drag-over');
            });
        });

        linkDropZone.addEventListener('drop', event => {
            const detected = droppedLinkData(event.dataTransfer);

            if (!detected) {
                showToast('Não consegui identificar uma URL válida no item arrastado.', true);
                return;
            }

            openLinkModal(detected);
            showToast('Link detectado. Escolha a categoria e salve.');
        });

        document.getElementById('importBrowserLinksBtn').addEventListener('click', openBrowserImportModal);
        document.getElementById('closeBrowserImportModalBtn').addEventListener('click', closeBrowserImportModal);
        document.getElementById('cancelBrowserImportBtn').addEventListener('click', closeBrowserImportModal);
        runBrowserImportBtn.addEventListener('click', runBrowserImport);
        browserImportList.addEventListener('change', event => {
            if (event.target.classList.contains('browser-import-choice')) {
                updateBrowserImportButton();
            }
        });
        browserImportModal.addEventListener('click', event => {
            if (event.target === browserImportModal) closeBrowserImportModal();
        });

        document.getElementById('openLinkModalBtn').addEventListener('click', () => openLinkModal());
        document.getElementById('closeLinkModalBtn').addEventListener('click', closeLinkModal);
        document.getElementById('cancelLinkBtn').addEventListener('click', closeLinkModal);
        linkForm.addEventListener('submit', saveLink);

        linksSearch.addEventListener('input', renderLinks);
        linksCategoryFilter.addEventListener('change', renderLinks);

        linkModal.addEventListener('click', event => {
            if (event.target === linkModal) closeLinkModal();
        });

        linksMobileMenuBtn.addEventListener('click', () => {
            linksSidebar.classList.add('open');
            linksSidebarOverlay.classList.remove('hidden');
        });

        linksSidebarOverlay.addEventListener('click', () => {
            linksSidebar.classList.remove('open');
            linksSidebarOverlay.classList.add('hidden');
        });

        document.addEventListener('keydown', event => {
            if (event.key === 'Escape') {
                closeLinkModal();
                closeBrowserImportModal();
                linksSidebar.classList.remove('open');
                linksSidebarOverlay.classList.add('hidden');
            }
        });

        loadLinks();
    </script>
</body>
</html>
"""



# ============================================================
# FRONTEND - APLICATIVOS LOCAIS
# ============================================================

LOCAL_APPS_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>Apps Locais • TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 76% 5%, rgba(44,196,219,.09), transparent 27rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }

        .glass-top {
            background: rgba(7,24,36,.90);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .sidebar-item {
            transition: border-color .18s ease, background .18s ease, color .18s ease;
        }

        .sidebar-item:hover,
        .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
        }

        .local-section {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(145deg, rgba(13,35,50,.93), rgba(8,25,37,.96));
        }

        .local-card {
            border: 1px solid rgba(129,180,204,.13);
            background: rgba(8,27,40,.80);
            transition: transform .16s ease, border-color .16s ease, background .16s ease;
        }

        .local-card:hover {
            transform: translateY(-2px);
            border-color: rgba(53,213,230,.28);
            background: rgba(12,39,54,.90);
        }

        .local-modal-backdrop {
            background: rgba(1,8,13,.80);
            backdrop-filter: blur(7px);
            -webkit-backdrop-filter: blur(7px);
        }

        @media (max-width: 1023px) {
            #localSidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }
            #localSidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased">
    <div id="localSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="localSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain" onerror="this.style.display='none'; this.nextElementSibling.classList.remove('hidden');">
            <div class="hidden items-center justify-center rounded-2xl border border-cyan-400/20 bg-cyan-400/10 p-4 text-3xl">⬡</div>
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▦</span><span class="font-semibold">APPS + USADOS</span>
            </a>

            <a href="/local-apps" class="sidebar-item active relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">▣</span><span class="flex-1 font-semibold">APPS LOCAIS</span>
            </a>

            <a href="/links" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">★</span><span class="font-semibold">LINKS FAVORITOS</span>
            </a>

            <a href="/workspace" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">◈</span><span class="font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>

            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span><span class="font-semibold">DIAGNÓSTICO WIN</span>
            </a>

            <a href="/" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="font-semibold">All tools</span>
            </a>
        </nav>

        <div class="border-t border-slate-700/20 p-3">
            <div class="rounded-xl border border-slate-700/20 bg-[#081b28] p-3">
                <div class="flex items-center gap-2">
                    <span class="h-2 w-2 rounded-full bg-emerald-400"></span>
                    <span class="text-xs font-semibold text-slate-300">Detecção local</span>
                </div>
                <p class="mt-1 pl-4 text-[10px] text-slate-600">Apps do computador</p>
            </div>
        </div>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex min-h-[72px] flex-wrap items-center gap-3 px-4 py-3 sm:px-6 xl:px-8">
                <button id="localMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden">☰</button>

                <div class="min-w-[190px] flex-1">
                    <h1 class="truncate text-base font-bold tracking-[.12em] text-slate-100">APPS LOCAIS</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Windows + Apple</p>
                </div>

                <div class="relative w-full sm:w-[300px]">
                    <span class="pointer-events-none absolute left-3 top-1/2 -translate-y-1/2 text-slate-600">⌕</span>
                    <input id="localSearch" type="search" placeholder="Buscar aplicativo local..." class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#0b2232]/85 pl-10 pr-3 text-sm text-slate-200 outline-none placeholder:text-slate-600 focus:border-cyan-400/35">
                </div>

                <select id="statusFilter" class="h-10 rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                    <option value="all">Todos</option>
                    <option value="installed">Instalados</option>
                    <option value="missing">Não encontrados</option>
                </select>

                <button id="addLocalAppBtn" class="h-10 rounded-xl border border-cyan-300/20 bg-cyan-400/10 px-4 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">
                    ＋ Adicionar App Local
                </button>

                <button id="refreshLocalApps" class="h-10 rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-4 text-xs font-bold text-slate-400 hover:text-cyan-200">
                    ↻ Atualizar
                </button>
            </div>
        </header>

        <main class="mx-auto max-w-[1400px] px-4 py-6 sm:px-6 xl:px-8">
            <section class="mb-5">
                <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Computador local</p>
                <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Aplicativos instalados</h2>
                <p class="mt-1 max-w-3xl text-sm text-slate-500">
                    O Hub verifica aplicativos conhecidos do Windows e da Apple. Os encontrados podem ser abertos diretamente por esta página.
                </p>
            </section>

            <div id="localLoading" class="py-16 text-center text-sm text-slate-600">Verificando aplicativos locais...</div>
            <div id="localGroups" class="hidden space-y-6"></div>

            <div id="localEmpty" class="hidden rounded-2xl border border-dashed border-slate-700/30 bg-[#081a27]/60 px-6 py-16 text-center">
                <div class="text-4xl">▣</div>
                <p class="mt-3 text-sm font-semibold text-slate-300">Nenhum aplicativo corresponde ao filtro.</p>
            </div>
        </main>
    </div>

    <!-- MODAL - ADICIONAR APP LOCAL -->
    <div id="localAppModal" class="local-modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-xl rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 class="text-xl font-bold text-slate-100">Adicionar aplicativo local</h3>
                    <p class="mt-1 text-xs text-slate-500">Cadastre um programa instalado que não aparece na detecção automática.</p>
                </div>
                <button id="closeLocalAppModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <form id="localAppForm" class="space-y-4">
                <div class="grid gap-4 sm:grid-cols-[1fr_150px]">
                    <div>
                        <label for="customLocalName" class="mb-1.5 block text-xs font-semibold text-slate-400">Nome</label>
                        <input id="customLocalName" type="text" maxlength="100" required placeholder="Ex.: LibreOffice" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                    </div>

                    <div>
                        <label for="customLocalGroup" class="mb-1.5 block text-xs font-semibold text-slate-400">Grupo</label>
                        <select id="customLocalGroup" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                            <option value="Windows">Windows</option>
                            <option value="Apple">Apple</option>
                        </select>
                    </div>
                </div>

                <div>
                    <label for="customLocalPath" class="mb-1.5 block text-xs font-semibold text-slate-400">Executável (.exe)</label>
                    <div class="flex gap-2">
                        <input id="customLocalPath" type="text" maxlength="1000" required placeholder="C:\\Program Files\\LibreOffice\\program\\soffice.exe" class="h-11 min-w-0 flex-1 rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                        <button id="browseLocalExeBtn" type="button" class="h-11 shrink-0 rounded-xl border border-slate-700/35 bg-[#0b2232] px-4 text-xs font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">
                            Procurar .EXE
                        </button>
                    </div>
                    <p class="mt-1.5 text-[10px] text-slate-600">O caminho é salvo apenas neste computador em local_apps_data.json.</p>
                </div>

                <div class="grid gap-4 sm:grid-cols-[100px_1fr]">
                    <div>
                        <label for="customLocalIcon" class="mb-1.5 block text-xs font-semibold text-slate-400">Ícone</label>
                        <input id="customLocalIcon" type="text" maxlength="16" value="🖥️" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-center text-xl text-slate-200 outline-none focus:border-cyan-400/35">
                    </div>

                    <div>
                        <label for="customLocalDescription" class="mb-1.5 block text-xs font-semibold text-slate-400">Descrição</label>
                        <input id="customLocalDescription" type="text" maxlength="240" placeholder="Ex.: Suíte de escritório" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                    </div>
                </div>

                <div id="customLocalError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:justify-end">
                    <button id="cancelLocalAppBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 hover:bg-slate-800/35 hover:text-slate-200">Cancelar</button>
                    <button id="saveLocalAppBtn" type="submit" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15 disabled:opacity-50">Salvar aplicativo</button>
                </div>
            </form>
        </div>
    </div>

    <div id="localToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const localGroups = document.getElementById('localGroups');
        const localLoading = document.getElementById('localLoading');
        const localEmpty = document.getElementById('localEmpty');
        const localSearch = document.getElementById('localSearch');
        const statusFilter = document.getElementById('statusFilter');
        const localToast = document.getElementById('localToast');
        const localSidebar = document.getElementById('localSidebar');
        const localSidebarOverlay = document.getElementById('localSidebarOverlay');

        const localAppModal = document.getElementById('localAppModal');
        const localAppForm = document.getElementById('localAppForm');
        const customLocalName = document.getElementById('customLocalName');
        const customLocalGroup = document.getElementById('customLocalGroup');
        const customLocalPath = document.getElementById('customLocalPath');
        const customLocalIcon = document.getElementById('customLocalIcon');
        const customLocalDescription = document.getElementById('customLocalDescription');
        const customLocalError = document.getElementById('customLocalError');
        const saveLocalAppBtn = document.getElementById('saveLocalAppBtn');

        let localApps = [];

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function normalize(value) {
            return String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase();
        }

        function showLocalToast(message, error = false) {
            localToast.textContent = message;
            localToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            localToast.classList.remove('hidden');
            clearTimeout(showLocalToast.timer);
            showLocalToast.timer = setTimeout(() => localToast.classList.add('hidden'), 2400);
        }

        function filteredLocalApps() {
            const query = normalize(localSearch.value.trim());
            const status = statusFilter.value;

            return localApps.filter(app => {
                if (status === 'installed' && !app.installed) return false;
                if (status === 'missing' && app.installed) return false;

                if (!query) return true;

                return [app.name, app.group, app.description]
                    .some(value => normalize(value).includes(query));
            });
        }

        function renderLocalApps() {
            const visible = filteredLocalApps();
            const groups = ['Windows', 'Apple'];

            localGroups.innerHTML = groups.map(group => {
                const items = visible.filter(app => app.group === group);
                if (!items.length) return '';

                const installedCount = items.filter(app => app.installed).length;

                return `
                    <section class="local-section rounded-2xl p-4 sm:p-5">
                        <div class="mb-4 flex items-center justify-between gap-3 border-b border-slate-700/20 pb-3">
                            <div>
                                <h3 class="text-base font-bold text-slate-200">${group}</h3>
                                <p class="mt-0.5 text-[11px] text-slate-600">${installedCount} encontrado${installedCount === 1 ? '' : 's'} nesta lista</p>
                            </div>
                            <span class="rounded-lg border border-slate-700/20 bg-[#061722] px-2.5 py-1 text-[10px] text-slate-500">${items.length} apps</span>
                        </div>

                        <div class="grid grid-cols-1 gap-3 sm:grid-cols-2 xl:grid-cols-3">
                            ${items.map(app => `
                                <article class="local-card flex min-h-[128px] items-center gap-4 rounded-xl p-4">
                                    <div class="grid h-14 w-14 shrink-0 place-items-center rounded-2xl border border-slate-700/25 bg-[#061722] text-3xl">${escapeHtml(app.icon)}</div>

                                    <div class="min-w-0 flex-1">
                                        <div class="flex items-start justify-between gap-2">
                                            <h4 class="truncate text-sm font-bold text-slate-200">${escapeHtml(app.name)}</h4>
                                            <span class="mt-1 h-2 w-2 shrink-0 rounded-full ${app.installed ? 'bg-emerald-400' : 'bg-slate-700'}"></span>
                                        </div>

                                        <p class="mt-1 line-clamp-2 text-[11px] leading-4 text-slate-600">${escapeHtml(app.description)}</p>

                                        <div class="mt-3 flex items-center justify-between gap-2">
                                            <span class="truncate text-[10px] ${app.installed ? 'text-emerald-400/80' : 'text-slate-700'}">
                                                ${app.installed ? '✓ ' + escapeHtml(app.source) : 'Não encontrado'}
                                            </span>

                                            <div class="flex items-center gap-1">
                                                ${app.custom ? `
                                                    <button
                                                        type="button"
                                                        data-delete-local-id="${escapeHtml(app.id)}"
                                                        data-delete-local-name="${escapeHtml(app.name)}"
                                                        class="rounded-lg px-2 py-1.5 text-[10px] text-slate-600 hover:bg-red-950/30 hover:text-red-300"
                                                        title="Excluir cadastro"
                                                    >✕</button>
                                                ` : ''}

                                                <button
                                                    type="button"
                                                    data-launch-id="${escapeHtml(app.id)}"
                                                    ${app.installed ? '' : 'disabled'}
                                                    class="rounded-lg border px-3 py-1.5 text-[10px] font-bold transition ${
                                                        app.installed
                                                            ? 'border-cyan-300/20 bg-cyan-400/10 text-cyan-200 hover:bg-cyan-400/15'
                                                            : 'cursor-not-allowed border-slate-700/20 bg-slate-900/20 text-slate-700'
                                                    }"
                                                >
                                                    ${app.installed ? 'Abrir' : 'Indisponível'}
                                                </button>
                                            </div>
                                        </div>
                                    </div>
                                </article>
                            `).join('')}
                        </div>
                    </section>
                `;
            }).join('');

            const hasVisible = visible.length > 0;
            localGroups.classList.toggle('hidden', !hasVisible);
            localEmpty.classList.toggle('hidden', hasVisible);
        }

        async function loadLocalApps() {
            localLoading.classList.remove('hidden');
            localGroups.classList.add('hidden');
            localEmpty.classList.add('hidden');

            try {
                const response = await fetch('/api/local-apps', { cache: 'no-store' });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível detectar os aplicativos.');
                }

                localApps = data;
                renderLocalApps();
            } catch (error) {
                showLocalToast(error.message || 'Erro na detecção de aplicativos.', true);
            } finally {
                localLoading.classList.add('hidden');
            }
        }

        function openLocalAppModal() {
            localAppForm.reset();
            customLocalGroup.value = 'Windows';
            customLocalIcon.value = '🖥️';
            customLocalError.classList.add('hidden');
            customLocalError.textContent = '';

            localAppModal.classList.remove('hidden');
            localAppModal.classList.add('flex');

            setTimeout(() => customLocalName.focus(), 50);
        }

        function closeLocalAppModal() {
            localAppModal.classList.add('hidden');
            localAppModal.classList.remove('flex');
            localAppForm.reset();
            customLocalError.classList.add('hidden');
        }

        async function browseLocalExecutable() {
            const button = document.getElementById('browseLocalExeBtn');
            const oldText = button.textContent;
            button.disabled = true;
            button.textContent = 'Abrindo...';

            try {
                const response = await fetch('/api/local-apps/browse-exe', {
                    method: 'POST'
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível abrir o seletor de arquivo.');
                }

                if (data.path) {
                    customLocalPath.value = data.path;

                    if (!customLocalName.value.trim() && data.suggested_name) {
                        customLocalName.value = data.suggested_name;
                    }
                }
            } catch (error) {
                showLocalToast(error.message || 'Erro ao selecionar executável.', true);
            } finally {
                button.disabled = false;
                button.textContent = oldText;
            }
        }

        async function createCustomLocalApp(event) {
            event.preventDefault();

            const payload = {
                name: customLocalName.value.trim(),
                group: customLocalGroup.value,
                executable_path: customLocalPath.value.trim(),
                icon: customLocalIcon.value.trim() || '🖥️',
                description: customLocalDescription.value.trim()
            };

            saveLocalAppBtn.disabled = true;
            saveLocalAppBtn.textContent = 'Salvando...';
            customLocalError.classList.add('hidden');

            try {
                const response = await fetch('/api/local-apps/custom', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const data = await response.json();

                if (!response.ok) {
                    const detail = typeof data.detail === 'string'
                        ? data.detail
                        : 'Não foi possível cadastrar o aplicativo.';
                    throw new Error(detail);
                }

                closeLocalAppModal();
                await loadLocalApps();
                showLocalToast(`“${data.name}” foi adicionado aos Apps Locais.`);
            } catch (error) {
                customLocalError.textContent = error.message || 'Erro ao salvar aplicativo local.';
                customLocalError.classList.remove('hidden');
            } finally {
                saveLocalAppBtn.disabled = false;
                saveLocalAppBtn.textContent = 'Salvar aplicativo';
            }
        }

        async function deleteCustomLocalApp(id, name) {
            if (!window.confirm(`Deseja remover “${name}” da lista de Apps Locais?`)) return;

            try {
                const response = await fetch(`/api/local-apps/custom/${encodeURIComponent(id)}`, {
                    method: 'DELETE'
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível excluir o cadastro.');
                }

                await loadLocalApps();
                showLocalToast(`“${name}” foi removido da lista.`);
            } catch (error) {
                showLocalToast(error.message || 'Erro ao excluir aplicativo local.', true);
            }
        }

        async function launchLocalApp(id) {
            try {
                const response = await fetch(`/api/local-apps/${encodeURIComponent(id)}/launch`, {
                    method: 'POST'
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível abrir o aplicativo.');
                }

                showLocalToast(data.message || 'Aplicativo aberto.');
            } catch (error) {
                showLocalToast(error.message || 'Erro ao abrir aplicativo.', true);
            }
        }

        localGroups.addEventListener('click', event => {
            const deleteButton = event.target.closest('[data-delete-local-id]');

            if (deleteButton) {
                deleteCustomLocalApp(
                    deleteButton.dataset.deleteLocalId,
                    deleteButton.dataset.deleteLocalName
                );
                return;
            }

            const button = event.target.closest('[data-launch-id]');
            if (!button || button.disabled) return;
            launchLocalApp(button.dataset.launchId);
        });

        localSearch.addEventListener('input', renderLocalApps);
        statusFilter.addEventListener('change', renderLocalApps);
        document.getElementById('refreshLocalApps').addEventListener('click', loadLocalApps);
        document.getElementById('addLocalAppBtn').addEventListener('click', openLocalAppModal);
        document.getElementById('closeLocalAppModalBtn').addEventListener('click', closeLocalAppModal);
        document.getElementById('cancelLocalAppBtn').addEventListener('click', closeLocalAppModal);
        document.getElementById('browseLocalExeBtn').addEventListener('click', browseLocalExecutable);
        localAppForm.addEventListener('submit', createCustomLocalApp);

        localAppModal.addEventListener('click', event => {
            if (event.target === localAppModal) closeLocalAppModal();
        });

        document.getElementById('localMenuBtn').addEventListener('click', () => {
            localSidebar.classList.add('open');
            localSidebarOverlay.classList.remove('hidden');
        });

        localSidebarOverlay.addEventListener('click', () => {
            localSidebar.classList.remove('open');
            localSidebarOverlay.classList.add('hidden');
        });

        loadLocalApps();
    </script>
</body>
</html>
"""



# ============================================================
# FRONTEND - DIAGNÓSTICO WINDOWS
# ============================================================

WINDOWS_DIAGNOSTICS_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>Diagnóstico Windows • TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 78% 4%, rgba(44,196,219,.09), transparent 27rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }

        .glass-top {
            background: rgba(7,24,36,.90);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .sidebar-item {
            transition: border-color .18s ease, background .18s ease, color .18s ease;
        }

        .sidebar-item:hover,
        .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
        }

        .diag-card {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(145deg, rgba(13,35,50,.93), rgba(8,25,37,.96));
            transition: border-color .18s ease, transform .18s ease;
        }

        .diag-card:hover {
            border-color: rgba(53,213,230,.24);
            transform: translateY(-2px);
        }

        .cleanup-panel {
            border: 1px solid rgba(79, 196, 230, .30);
            background: linear-gradient(145deg, rgba(7,31,44,.98), rgba(5,24,35,.98));
            box-shadow: inset 0 1px rgba(255,255,255,.02);
        }

        .cleanup-header {
            border-bottom: 1px solid rgba(79,196,230,.25);
        }

        .cleanup-row {
            border-bottom: 1px solid rgba(129,180,204,.13);
        }

        .cleanup-row:last-child {
            border-bottom: 0;
        }

        .cleanup-checkbox {
            accent-color: #67e8f9;
        }

        @media (max-width: 1023px) {
            #diagSidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }
            #diagSidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased">
    <div id="diagSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="diagSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain" onerror="this.style.display='none'; this.nextElementSibling.classList.remove('hidden');">
            <div class="hidden items-center justify-center rounded-2xl border border-cyan-400/20 bg-cyan-400/10 p-4 text-3xl">⬡</div>
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▦</span><span class="font-semibold">APPS + USADOS</span>
            </a>

            <a href="/local-apps" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▣</span><span class="font-semibold">APPS LOCAIS</span>
            </a>

            <a href="/links" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">★</span><span class="font-semibold">LINKS FAVORITOS</span>
            </a>

            <a href="/workspace" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">◈</span><span class="font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>

            <a href="/windows-diagnostics" class="sidebar-item active relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">🩺</span><span class="flex-1 font-semibold">DIAGNÓSTICO WIN</span>
            </a>

            <a href="/" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="font-semibold">All tools</span>
            </a>
        </nav>

        <div class="border-t border-slate-700/20 p-3">
            <div class="rounded-xl border border-slate-700/20 bg-[#081b28] p-3">
                <div class="flex items-center gap-2">
                    <span class="h-2 w-2 rounded-full bg-emerald-400"></span>
                    <span class="text-xs font-semibold text-slate-300">Ferramentas nativas</span>
                </div>
                <p class="mt-1 pl-4 text-[10px] text-slate-600">Windows 10 / 11</p>
            </div>
        </div>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex h-[72px] items-center gap-3 px-4 sm:px-6 xl:px-8">
                <button id="diagMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden">☰</button>

                <div class="min-w-0 flex-1">
                    <h1 class="truncate text-base font-bold tracking-[.12em] text-slate-100">DIAGNÓSTICO WINDOWS</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Manutenção e diagnóstico nativo</p>
                </div>

                <button id="refreshDiagStatus" class="h-10 rounded-xl border border-slate-700/30 bg-[#0b2232]/85 px-4 text-xs font-bold text-slate-400 hover:text-cyan-200">
                    ↻ Atualizar status
                </button>
            </div>
        </header>

        <main class="mx-auto max-w-[1350px] px-4 py-6 sm:px-6 xl:px-8">
            <section class="mb-6">
                <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Windows</p>
                <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Ferramentas de diagnóstico</h2>
                <p class="mt-1 max-w-3xl text-sm text-slate-500">
                    Atalhos seguros para recursos nativos do Windows. Ações que podem reiniciar o computador ou excluir dados continuam exigindo confirmação nas próprias telas do Windows ou do navegador.
                </p>
            </section>

            <div class="grid gap-5 xl:grid-cols-3">
                <!-- LIMPEZA DE DISCO -->
                <section class="cleanup-panel overflow-hidden rounded-2xl xl:col-span-1">
                    <div class="cleanup-header flex items-center justify-between px-4 py-3">
                        <div class="text-[11px] font-black uppercase tracking-[.17em] text-cyan-100">Limpeza de Disco</div>
                        <div class="text-[9px] font-mono text-cyan-500/70">02</div>
                    </div>

                    <div class="p-4">
                        <p class="mb-3 text-[10px] text-cyan-200/65">Selecione as áreas para analisar.</p>

                        <label class="mb-2 flex cursor-pointer items-center gap-2 rounded-lg border border-cyan-300/55 bg-cyan-400/[.04] px-3 py-2.5 text-[10px] font-bold text-cyan-100">
                            <input id="cleanupSelectAll" type="checkbox" checked class="cleanup-checkbox h-3.5 w-3.5">
                            <span>Marcar todas as opções</span>
                        </label>

                        <div id="cleanupRows">
                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="windows_temp" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-cyan-400"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Windows Temp <span class="font-normal text-cyan-100/70">(Prefetch, %TEMP%, %TMP%)</span></span>
                                <span data-cleanup-size="windows_temp" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>

                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="recycle_bin" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-emerald-400"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Lixeira</span>
                                <span data-cleanup-size="recycle_bin" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>

                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="windows_cache" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-violet-400"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Caches/diagnósticos do Windows</span>
                                <span data-cleanup-size="windows_cache" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>

                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="browser_cache" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-sky-400"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Cache dos navegadores</span>
                                <span data-cleanup-size="browser_cache" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>

                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="browser_history" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-fuchsia-400"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Histórico de navegação</span>
                                <span data-cleanup-size="browser_history" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>

                            <label class="cleanup-row flex cursor-pointer items-center gap-2 py-2.5">
                                <input type="checkbox" value="browser_sessions" checked class="cleanup-area cleanup-checkbox h-3.5 w-3.5">
                                <span class="h-1.5 w-1.5 rounded-sm bg-amber-300"></span>
                                <span class="min-w-0 flex-1 text-[11px] font-semibold text-cyan-50">Cookies e sessões</span>
                                <span data-cleanup-size="browser_sessions" class="shrink-0 font-mono text-[11px] font-bold text-cyan-200">—</span>
                            </label>
                        </div>

                        <div class="mt-3 flex items-end justify-between gap-3 border-t border-cyan-400/15 pt-3">
                            <div>
                                <div class="text-[9px] text-cyan-100/55">Elegível nas áreas analisadas</div>
                                <div id="cleanupTotal" class="mt-1 font-mono text-lg text-cyan-200">0 B</div>
                            </div>
                            <div id="cleanupStatus" class="text-right text-[9px] text-cyan-100/45">Ainda não analisado</div>
                        </div>

                        <div class="mt-3 grid grid-cols-2 gap-2">
                            <button id="scanCleanupBtn" type="button" class="h-10 rounded-lg border border-amber-500/45 bg-amber-500/[.06] text-[10px] font-bold text-amber-200 hover:bg-amber-500/[.10]">
                                ⛶ &nbsp; Escanear agora
                            </button>

                            <button id="cleanCleanupBtn" type="button" class="h-10 rounded-lg border border-cyan-300/35 bg-cyan-400/[.06] text-[10px] font-bold text-cyan-200 hover:bg-cyan-400/[.12]">
                                Limpar selecionados
                            </button>
                        </div>

                        <p class="mt-3 text-[9px] leading-4 text-cyan-100/40">
                            Feche os navegadores antes de limpar cache, histórico, cookies ou sessões. Arquivos em uso são ignorados.
                        </p>
                    </div>
                </section>

                <!-- PERFMON -->
                <section class="diag-card rounded-2xl p-5">
                    <div class="flex items-start gap-4">
                        <div class="grid h-12 w-12 shrink-0 place-items-center rounded-xl border border-cyan-400/15 bg-cyan-400/5 text-2xl">📈</div>
                        <div class="min-w-0">
                            <div class="text-xs font-semibold uppercase tracking-[.12em] text-cyan-400/55">2 • Desempenho</div>
                            <h3 class="mt-1 text-base font-bold text-slate-200">Monitor de Recursos</h3>
                            <p class="mt-1 text-xs leading-5 text-slate-500">Executa o relatório de diagnóstico de desempenho do Windows.</p>
                        </div>
                    </div>

                    <div class="mt-5 rounded-xl border border-slate-700/20 bg-[#061722] px-3 py-3 font-mono text-xs text-slate-400">perfmon /report</div>
                    <p class="mt-2 text-[10px] leading-4 text-slate-600">O Windows solicitará permissão de administrador (UAC) e normalmente coletará dados por cerca de 60 segundos.</p>

                    <button data-diag-action="performance" class="mt-4 h-10 w-full rounded-xl border border-cyan-300/20 bg-cyan-400/10 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">
                        Gerar relatório
                    </button>
                </section>

                <!-- MEMÓRIA -->
                <section class="diag-card rounded-2xl p-5">
                    <div class="flex items-start gap-4">
                        <div class="grid h-12 w-12 shrink-0 place-items-center rounded-xl border border-cyan-400/15 bg-cyan-400/5 text-2xl">🧠</div>
                        <div class="min-w-0">
                            <div class="text-xs font-semibold uppercase tracking-[.12em] text-cyan-400/55">3 • Memória RAM</div>
                            <h3 class="mt-1 text-base font-bold text-slate-200">Diagnóstico de Memória</h3>
                            <p class="mt-1 text-xs leading-5 text-slate-500">Abre a ferramenta oficial para verificar possíveis erros de memória.</p>
                        </div>
                    </div>

                    <div class="mt-5 rounded-xl border border-slate-700/20 bg-[#061722] px-3 py-3 font-mono text-xs text-slate-400">mdsched.exe</div>
                    <p class="mt-2 text-[10px] leading-4 text-amber-300/65">O reinício só ocorrerá se você escolher essa opção na janela do Windows.</p>

                    <button data-diag-action="memory" class="mt-4 h-10 w-full rounded-xl border border-cyan-300/20 bg-cyan-400/10 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">
                        Abrir diagnóstico
                    </button>
                </section>
            </div>

        </main>
    </div>

    <div id="diagToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const diagToast = document.getElementById('diagToast');
        const diagSidebar = document.getElementById('diagSidebar');
        const diagSidebarOverlay = document.getElementById('diagSidebarOverlay');

        const cleanupSelectAll = document.getElementById('cleanupSelectAll');
        const cleanupTotal = document.getElementById('cleanupTotal');
        const cleanupStatus = document.getElementById('cleanupStatus');
        const scanCleanupBtn = document.getElementById('scanCleanupBtn');
        const cleanCleanupBtn = document.getElementById('cleanCleanupBtn');

        let cleanupScan = new Map();

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function showDiagToast(message, error = false) {
            diagToast.textContent = message;
            diagToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            diagToast.classList.remove('hidden');
            clearTimeout(showDiagToast.timer);
            showDiagToast.timer = setTimeout(() => diagToast.classList.add('hidden'), 2800);
        }

        function cleanupCheckboxes() {
            return [...document.querySelectorAll('.cleanup-area')];
        }

        function selectedCleanupAreas() {
            return cleanupCheckboxes()
                .filter(input => input.checked)
                .map(input => input.value);
        }

        function formatBytes(bytes) {
            const value = Math.max(0, Number(bytes) || 0);

            if (value < 1024) return `${value} B`;

            const units = ['KB', 'MB', 'GB', 'TB'];
            let current = value;
            let unit = 'B';

            for (const nextUnit of units) {
                current /= 1024;
                unit = nextUnit;
                if (current < 1024) break;
            }

            const decimals = current >= 100 ? 0 : current >= 10 ? 1 : 2;
            return `${current.toFixed(decimals).replace('.', ',')} ${unit}`;
        }

        function updateCleanupTotal() {
            const selected = new Set(selectedCleanupAreas());
            let total = 0;

            for (const [id, bytes] of cleanupScan.entries()) {
                if (selected.has(id)) total += Number(bytes) || 0;
            }

            cleanupTotal.textContent = formatBytes(total);

            const boxes = cleanupCheckboxes();
            cleanupSelectAll.checked = boxes.length > 0 && boxes.every(box => box.checked);
            cleanupSelectAll.indeterminate = boxes.some(box => box.checked) && !boxes.every(box => box.checked);
        }

        function setCleanupBusy(busy, mode = '') {
            scanCleanupBtn.disabled = busy;
            cleanCleanupBtn.disabled = busy;

            scanCleanupBtn.classList.toggle('opacity-50', busy);
            cleanCleanupBtn.classList.toggle('opacity-50', busy);

            if (!busy) {
                scanCleanupBtn.innerHTML = '⛶ &nbsp; Escanear agora';
                cleanCleanupBtn.textContent = 'Limpar selecionados';
                return;
            }

            if (mode === 'scan') {
                scanCleanupBtn.textContent = 'Analisando...';
            } else if (mode === 'clean') {
                cleanCleanupBtn.textContent = 'Limpando...';
            }
        }

        async function scanDiskCleanup() {
            const areas = selectedCleanupAreas();

            if (!areas.length) {
                showDiagToast('Selecione ao menos uma área para analisar.', true);
                return;
            }

            setCleanupBusy(true, 'scan');
            cleanupStatus.textContent = 'Analisando...';

            try {
                const response = await fetch('/api/windows-diagnostics/disk-cleanup/scan', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ areas, confirmed: false })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível analisar as áreas.');
                }

                cleanupScan = new Map();

                for (const area of data.areas || []) {
                    cleanupScan.set(area.id, Number(area.bytes) || 0);

                    const sizeEl = document.querySelector(`[data-cleanup-size="${area.id}"]`);
                    if (sizeEl) sizeEl.textContent = formatBytes(area.bytes);
                }

                updateCleanupTotal();
                cleanupStatus.textContent = `${areas.length} área(s) analisada(s)`;
                showDiagToast(`Varredura concluída: ${formatBytes(data.total_bytes)} elegíveis.`);
            } catch (error) {
                cleanupStatus.textContent = 'Falha na análise';
                showDiagToast(error.message || 'Erro na varredura.', true);
            } finally {
                setCleanupBusy(false);
            }
        }

        async function cleanSelectedAreas() {
            const areas = selectedCleanupAreas();

            if (!areas.length) {
                showDiagToast('Selecione ao menos uma área para limpar.', true);
                return;
            }

            const clearsBrowserData = areas.some(id =>
                ['browser_history', 'browser_sessions'].includes(id)
            );

            const warning = clearsBrowserData
                ? 'Os itens selecionados incluem histórico, cookies ou sessões. Isso pode desconectar contas dos navegadores. Feche os navegadores antes de continuar. Deseja limpar as áreas selecionadas?'
                : 'Deseja limpar as áreas selecionadas? Arquivos em uso serão ignorados.';

            if (!window.confirm(warning)) return;

            setCleanupBusy(true, 'clean');
            cleanupStatus.textContent = 'Limpando...';

            try {
                const response = await fetch('/api/windows-diagnostics/disk-cleanup/clean', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ areas, confirmed: true })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível concluir a limpeza.');
                }

                showDiagToast(
                    `Limpeza concluída: ${formatBytes(data.removed_bytes)} removidos` +
                    (data.skipped_items ? ` • ${data.skipped_items} item(ns) em uso/ignorados` : '')
                );

                cleanupStatus.textContent = 'Limpeza concluída';
                await scanDiskCleanup();
            } catch (error) {
                cleanupStatus.textContent = 'Falha na limpeza';
                showDiagToast(error.message || 'Erro na limpeza.', true);
            } finally {
                setCleanupBusy(false);
            }
        }

        async function loadDiagStatus() {
            try {
                const response = await fetch('/api/windows-diagnostics/status', { cache: 'no-store' });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível consultar o Windows.');
                }

                if (!data.windows) {
                    showDiagToast('Esta área foi projetada para Windows.', true);
                }
            } catch (error) {
                showDiagToast(error.message || 'Erro ao carregar diagnóstico.', true);
            }
        }

        async function runAction(action) {
            const endpointMap = {
                performance: '/api/windows-diagnostics/performance-report',
                memory: '/api/windows-diagnostics/memory-diagnostic'
            };

            const endpoint = endpointMap[action];
            if (!endpoint) return;

            try {
                const response = await fetch(endpoint, { method: 'POST' });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível executar a ferramenta.');
                }

                showDiagToast(data.message || 'Ferramenta iniciada.');
            } catch (error) {
                showDiagToast(error.message || 'Erro ao executar ferramenta.', true);
            }
        }

        document.addEventListener('click', event => {
            const actionButton = event.target.closest('[data-diag-action]');
            if (actionButton) {
                runAction(actionButton.dataset.diagAction);
            }
        });

        cleanupSelectAll.addEventListener('change', () => {
            cleanupCheckboxes().forEach(box => {
                box.checked = cleanupSelectAll.checked;
            });
            updateCleanupTotal();
        });

        document.getElementById('cleanupRows').addEventListener('change', event => {
            if (event.target.classList.contains('cleanup-area')) {
                updateCleanupTotal();
            }
        });

        scanCleanupBtn.addEventListener('click', scanDiskCleanup);
        cleanCleanupBtn.addEventListener('click', cleanSelectedAreas);

        document.getElementById('refreshDiagStatus').addEventListener('click', loadDiagStatus);

        document.getElementById('diagMenuBtn').addEventListener('click', () => {
            diagSidebar.classList.add('open');
            diagSidebarOverlay.classList.remove('hidden');
        });

        diagSidebarOverlay.addEventListener('click', () => {
            diagSidebar.classList.remove('open');
            diagSidebarOverlay.classList.add('hidden');
        });

        loadDiagStatus();
    </script>
</body>
</html>
"""



# ============================================================
# FRONTEND - PERSONALIZAR ÁREA / PROJETOS
# ============================================================

WORKSPACE_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>Personalizar Área • TECH TOOL HUB</title>

    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 76% 5%, rgba(44,196,219,.09), transparent 27rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }

        .glass-top {
            background: rgba(7,24,36,.90);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .sidebar-item {
            transition: border-color .18s ease, background .18s ease, color .18s ease;
        }

        .sidebar-item:hover,
        .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
        }

        .panel {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(145deg, rgba(13,35,50,.93), rgba(8,25,37,.96));
        }

        .explorer-shell {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(180deg, rgba(11,31,44,.96), rgba(7,24,36,.96));
        }

        .explorer-nav-item {
            transition: background .14s ease, color .14s ease, border-color .14s ease;
        }

        .explorer-nav-item:hover,
        .explorer-nav-item.active {
            background: rgba(35,135,162,.13);
            color: #bff7ff;
            border-color: rgba(53,213,230,.17);
        }

        .explorer-row {
            border-bottom: 1px solid rgba(129,180,204,.08);
            transition: background .12s ease, box-shadow .12s ease;
        }

        .explorer-row:hover { background: rgba(34,93,119,.11); }
        .explorer-row.selected {
            background: rgba(28,141,169,.14);
            box-shadow: inset 3px 0 rgba(53,213,230,.75);
        }

        .explorer-col-title {
            font-size: 10px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: .08em;
            color: rgb(71 85 105);
        }

        .folder-icon {
            position: relative;
            width: 28px;
            height: 20px;
            border-radius: 4px 5px 5px 5px;
            background: linear-gradient(180deg, #e4b44e, #c78a2b);
            box-shadow: inset 0 -2px rgba(98,58,12,.18);
        }

        .folder-icon::before {
            content: '';
            position: absolute;
            left: 2px;
            top: -5px;
            width: 12px;
            height: 7px;
            border-radius: 4px 4px 0 0;
            background: #efc260;
        }

        @media (max-width: 767px) {
            .explorer-table-head { display: none; }
            .explorer-row {
                display: block !important;
                padding: 14px;
                border: 1px solid rgba(129,180,204,.10);
                border-radius: 14px;
                margin-bottom: 8px;
            }
            .explorer-hide-mobile { display: none; }
        }

        .project-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(7,27,40,.78);
            transition: border-color .16s ease, background .16s ease, transform .16s ease;
        }

        .project-card:hover,
        .project-card.active {
            border-color: rgba(53,213,230,.38);
            background: rgba(13,43,58,.90);
        }

        .project-card.active {
            box-shadow: inset 3px 0 #35d5e6;
        }

        .metric-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(5,22,33,.64);
        }

        .phase-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(6,25,37,.72);
            transition: border-color .16s ease, transform .16s ease;
        }

        .phase-card:hover {
            border-color: rgba(53,213,230,.28);
            transform: translateY(-1px);
        }

        .phase-progress-track {
            height: 6px;
            border-radius: 999px;
            overflow: hidden;
            background: rgba(30,52,66,.75);
        }

        .phase-progress-fill,
        .overall-progress-fill {
            height: 100%;
            border-radius: inherit;
            background: linear-gradient(90deg, rgba(34,211,238,.78), rgba(59,130,246,.88));
            transition: width .25s ease;
        }

        .resource-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(7,27,40,.74);
            transition: transform .15s ease, border-color .15s ease, opacity .15s ease;
        }

        .resource-card[draggable="true"] { cursor: grab; }
        .resource-card[draggable="true"]:active { cursor: grabbing; }

        .resource-card:hover {
            border-color: rgba(53,213,230,.28);
            transform: translateY(-1px);
        }

        .resource-card.dragging { opacity: .35; }

        .project-drop-zone {
            min-height: 176px;
            border: 1px dashed rgba(53,213,230,.28);
            background: rgba(4,20,30,.36);
            transition: border-color .16s ease, background .16s ease, box-shadow .16s ease;
        }

        .project-drop-zone.drag-over {
            border-color: rgba(103,232,249,.75);
            background: rgba(8,145,178,.08);
            box-shadow: 0 0 0 2px rgba(53,213,230,.08), 0 0 28px rgba(53,213,230,.08);
        }

        .status-planejado { color: #93c5fd; background: rgba(37,99,235,.12); border-color: rgba(96,165,250,.22); }
        .status-em_andamento { color: #67e8f9; background: rgba(8,145,178,.12); border-color: rgba(34,211,238,.22); }
        .status-pausado { color: #fcd34d; background: rgba(180,83,9,.12); border-color: rgba(251,191,36,.22); }
        .status-concluido { color: #6ee7b7; background: rgba(5,150,105,.12); border-color: rgba(52,211,153,.22); }
        .status-cancelado { color: #fca5a5; background: rgba(185,28,28,.12); border-color: rgba(248,113,113,.22); }

        .modal-backdrop {
            background: rgba(1,8,13,.82);
            backdrop-filter: blur(7px);
            -webkit-backdrop-filter: blur(7px);
        }

        @media (max-width: 1023px) {
            #workspaceSidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }
            #workspaceSidebar.open { transform: translateX(0); }
        }
    </style>
</head>

<body class="text-slate-100 antialiased">
    <div id="workspaceSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="workspaceSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain">
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▦</span><span class="font-semibold">APPS + USADOS</span>
            </a>
            <a href="/local-apps" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▣</span><span class="font-semibold">APPS LOCAIS</span>
            </a>
            <a href="/links" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">★</span><span class="font-semibold">LINKS FAVORITOS</span>
            </a>
            <a href="/workspace" class="sidebar-item active relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">◈</span><span class="font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">⑂</span><span class="flex-1 font-semibold">GIT &amp; GITHUB</span>
            </a>
            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span><span class="font-semibold">DIAGNÓSTICO WIN</span>
            </a>
            <a href="/" class="sidebar-item relative mt-1 flex w-full items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="font-semibold">All tools</span>
            </a>
        </nav>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex min-h-[72px] flex-wrap items-center gap-3 px-4 py-3 sm:px-6 xl:px-8">
                <button id="workspaceMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden">☰</button>

                <div class="min-w-[210px] flex-1">
                    <h1 class="truncate text-base font-bold tracking-[.12em] text-slate-100">PERSONALIZAR ÁREA</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Projetos, clientes e ferramentas</p>
                </div>

                <button id="newProjectBtn" class="h-10 rounded-xl border border-cyan-300/20 bg-cyan-400/10 px-4 text-sm font-bold text-cyan-200 hover:bg-cyan-400/15">
                    ＋ Novo Projeto
                </button>
            </div>
        </header>

        <main class="mx-auto max-w-[1500px] px-4 py-6 sm:px-6 xl:px-8">
            <section class="mb-5">
                <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Workspace Explorer</p>
                <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Projetos em D:\python\CHATGPT</h2>
                <p class="mt-1 max-w-4xl text-sm text-slate-500">
                    As pastas reais são a fonte da lista. O cadastro do TECH TOOL HUB adiciona status, fases, progresso e ferramentas sem mover ou renomear seus diretórios.
                </p>
            </section>

            <section class="explorer-shell mb-5 overflow-hidden rounded-2xl">
                <div class="border-b border-slate-700/20 px-4 py-3">
                    <div class="flex flex-col gap-3 lg:flex-row lg:items-center">
                        <div class="flex min-w-0 flex-1 items-center gap-2">
                            <button id="workspaceExplorerRefreshBtn" type="button" class="grid h-9 w-9 shrink-0 place-items-center rounded-lg border border-slate-700/25 bg-[#071722] text-slate-400 hover:text-cyan-200" title="Atualizar">↻</button>
                            <div class="flex min-w-0 flex-1 items-center gap-2 rounded-lg border border-slate-700/25 bg-[#061722] px-3 py-2">
                                <span class="text-amber-300">📁</span>
                                <span id="workspaceExplorerRoot" class="truncate font-mono text-[11px] text-slate-400">D:\python\CHATGPT</span>
                            </div>
                            <button id="workspaceExplorerNewBtn" class="h-9 shrink-0 rounded-lg border border-cyan-400/15 bg-cyan-400/5 px-3 text-[11px] font-bold text-cyan-300">＋ Cadastro</button>
                        </div>
                        <div class="relative w-full lg:w-[330px]">
                            <span class="pointer-events-none absolute left-3 top-2.5 text-slate-600">⌕</span>
                            <input id="workspaceExplorerSearch" placeholder="Pesquisar projetos..." class="h-9 w-full rounded-lg border border-slate-700/25 bg-[#061722] pl-8 pr-3 text-xs text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/30">
                        </div>
                    </div>
                </div>

                <div class="grid min-h-[360px] lg:grid-cols-[210px_minmax(0,1fr)]">
                    <aside class="border-b border-slate-700/15 bg-[#061722]/60 p-3 lg:border-b-0 lg:border-r">
                        <div class="mb-2 px-2 text-[9px] font-bold uppercase tracking-[.14em] text-slate-700">Acesso rápido</div>
                        <button data-workspace-filter="all" class="explorer-nav-item active flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span>📂</span><span class="flex-1">Todos os projetos</span><span id="workspaceCountAll" class="text-[10px] text-slate-600">0</span></button>
                        <button data-workspace-filter="em_andamento" class="explorer-nav-item mt-1 flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span class="h-2 w-2 rounded-full bg-cyan-400"></span><span class="flex-1">Em andamento</span><span id="workspaceCountActive" class="text-[10px] text-slate-600">0</span></button>
                        <button data-workspace-filter="pausado" class="explorer-nav-item mt-1 flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span class="h-2 w-2 rounded-full bg-amber-300"></span><span class="flex-1">Pausados</span><span id="workspaceCountPaused" class="text-[10px] text-slate-600">0</span></button>
                        <button data-workspace-filter="concluido" class="explorer-nav-item mt-1 flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span class="h-2 w-2 rounded-full bg-emerald-400"></span><span class="flex-1">Concluídos</span><span id="workspaceCountDone" class="text-[10px] text-slate-600">0</span></button>
                        <button data-workspace-filter="nao_classificado" class="explorer-nav-item mt-1 flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span class="h-2 w-2 rounded-full bg-slate-600"></span><span class="flex-1">Sem cadastro</span><span id="workspaceCountUnregistered" class="text-[10px] text-slate-600">0</span></button>
                        <button data-workspace-filter="git" class="explorer-nav-item mt-1 flex w-full items-center gap-2 rounded-lg border border-transparent px-3 py-2 text-left text-xs text-slate-400"><span>⑂</span><span class="flex-1">Repositórios Git</span><span id="workspaceCountGit" class="text-[10px] text-slate-600">0</span></button>
                        <div class="mt-5 rounded-xl border border-emerald-400/10 bg-emerald-400/[.025] p-3 text-[10px] leading-4 text-slate-600"><strong class="text-emerald-300/80">Privacidade:</strong> a raiz e os caminhos escolhidos no Git ficam somente na sessão.</div>
                    </aside>

                    <section class="min-w-0">
                        <div class="explorer-table-head grid grid-cols-[minmax(260px,1.7fr)_150px_110px_90px_150px_130px] border-b border-slate-700/15 bg-[#071722]/50 px-4 py-2">
                            <div class="explorer-col-title">Nome</div><div class="explorer-col-title">Status</div><div class="explorer-col-title">Progresso</div><div class="explorer-col-title">Git</div><div class="explorer-col-title">Modificado</div><div class="explorer-col-title text-right">Ações</div>
                        </div>
                        <div id="workspaceExplorerList" class="max-h-[560px] overflow-y-auto"></div>
                        <div id="workspaceExplorerEmpty" class="hidden px-5 py-14 text-center"><div class="text-4xl">📂</div><p class="mt-2 text-xs font-semibold text-slate-400">Nenhuma pasta encontrada</p><p class="mt-1 text-[10px] text-slate-600">Verifique o filtro selecionado.</p></div>
                    </section>
                </div>
                <div class="flex flex-wrap items-center justify-between gap-2 border-t border-slate-700/15 bg-[#061722]/55 px-4 py-2"><span id="workspaceExplorerCount" class="text-[10px] text-slate-600">0 projetos</span><span id="workspaceExplorerSelection" class="max-w-full truncate text-[10px] text-slate-600">Nenhum projeto selecionado</span></div>
            </section>

            <div class="space-y-5">
                <!-- LISTA DE PROJETOS -->
                <aside class="hidden">
                    <div class="mb-3 flex items-center justify-between gap-3">
                        <div>
                            <h3 class="text-sm font-bold text-slate-200">Projetos</h3>
                            <p id="projectsCount" class="mt-0.5 text-[10px] text-slate-600">0 projetos</p>
                        </div>
                        <button id="newProjectMiniBtn" class="grid h-8 w-8 place-items-center rounded-lg border border-cyan-400/15 bg-cyan-400/5 text-cyan-300" title="Novo projeto">＋</button>
                    </div>

                    <div id="projectsList" class="space-y-2"></div>

                    <div id="projectsEmpty" class="hidden rounded-xl border border-dashed border-slate-700/30 px-4 py-10 text-center">
                        <div class="text-3xl">◈</div>
                        <p class="mt-2 text-xs font-semibold text-slate-400">Nenhum projeto criado</p>
                    </div>
                </aside>

                <!-- ÁREA DO PROJETO -->
                <section id="projectArea" class="hidden space-y-5">
                    <div class="panel rounded-2xl p-5">
                        <div class="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
                            <div class="min-w-0">
                                <div class="flex flex-wrap items-center gap-2">
                                    <h3 id="projectTitle" class="truncate text-xl font-bold text-slate-100"></h3>
                                    <span id="projectStatusBadge" class="rounded-lg border px-2.5 py-1 text-[10px] font-bold uppercase tracking-[.08em]"></span>
                                </div>

                                <div class="mt-3 flex flex-wrap gap-x-6 gap-y-2 text-xs text-slate-500">
                                    <span>Cliente: <strong id="projectClient" class="font-semibold text-slate-300">—</strong></span>
                                    <span>Início: <strong id="projectStart" class="font-semibold text-slate-300">—</strong></span>
                                    <span>Final: <strong id="projectEnd" class="font-semibold text-slate-300">—</strong></span>
                                </div>

                                <p id="projectNotes" class="mt-3 hidden max-w-3xl text-xs leading-5 text-slate-500"></p>
                            </div>

                            <div class="flex shrink-0 gap-2">
                                <button id="editProjectBtn" class="h-9 rounded-lg border border-slate-700/30 px-3 text-xs font-semibold text-slate-400 hover:border-cyan-400/25 hover:text-cyan-200">✎ Editar</button>
                                <button id="deleteProjectBtn" class="h-9 rounded-lg border border-red-900/30 px-3 text-xs font-semibold text-red-300/70 hover:bg-red-950/30 hover:text-red-200">Excluir</button>
                            </div>
                        </div>
                    </div>

                    <!-- DASHBOARD DO PROJETO -->
                    <section class="panel rounded-2xl p-5">
                        <div class="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                            <div>
                                <h3 class="text-sm font-bold text-slate-200">Dashboard do projeto</h3>
                                <p class="mt-0.5 text-[10px] text-slate-600">Fases, progresso e situação atual</p>
                            </div>
                            <button id="newPhaseBtn" type="button" class="h-9 rounded-lg border border-cyan-400/20 bg-cyan-400/5 px-3 text-xs font-bold text-cyan-200 hover:bg-cyan-400/10">
                                ＋ Nova fase
                            </button>
                        </div>

                        <div class="mt-4 grid grid-cols-2 gap-3 lg:grid-cols-4">
                            <div class="metric-card rounded-xl p-3">
                                <div class="text-[9px] font-semibold uppercase tracking-[.12em] text-slate-600">Progresso</div>
                                <div id="dashboardProgress" class="mt-1 text-2xl font-black text-cyan-200">0%</div>
                            </div>
                            <div class="metric-card rounded-xl p-3">
                                <div class="text-[9px] font-semibold uppercase tracking-[.12em] text-slate-600">Fases concluídas</div>
                                <div id="dashboardPhases" class="mt-1 text-2xl font-black text-slate-200">0/0</div>
                            </div>
                            <div class="metric-card rounded-xl p-3">
                                <div class="text-[9px] font-semibold uppercase tracking-[.12em] text-slate-600">Prazo</div>
                                <div id="dashboardDeadline" class="mt-1 truncate text-base font-bold text-slate-200">—</div>
                            </div>
                            <div class="metric-card rounded-xl p-3">
                                <div class="text-[9px] font-semibold uppercase tracking-[.12em] text-slate-600">Ferramentas</div>
                                <div id="dashboardResources" class="mt-1 text-2xl font-black text-slate-200">0</div>
                            </div>
                        </div>

                        <div class="mt-4">
                            <div class="mb-1.5 flex items-center justify-between gap-3">
                                <span class="text-[10px] font-semibold text-slate-500">Andamento geral</span>
                                <span id="dashboardProgressLabel" class="text-[10px] font-bold text-cyan-300">0%</span>
                            </div>
                            <div class="h-2 overflow-hidden rounded-full bg-[#061722]">
                                <div id="dashboardProgressBar" class="overall-progress-fill" style="width:0%"></div>
                            </div>
                        </div>

                        <div class="mt-5 flex items-center justify-between gap-3">
                            <div>
                                <h4 class="text-xs font-bold text-slate-300">Fases do projeto</h4>
                                <p class="mt-0.5 text-[10px] text-slate-600">O progresso geral é a média das fases.</p>
                            </div>
                            <span id="phaseCountBadge" class="rounded-lg border border-slate-700/25 bg-[#061722] px-2 py-1 text-[10px] text-slate-500">0 fases</span>
                        </div>

                        <div id="phaseTimeline" class="mt-3 grid gap-3 md:grid-cols-2 xl:grid-cols-4"></div>

                        <div id="phaseEmpty" class="mt-3 hidden rounded-xl border border-dashed border-slate-700/30 px-4 py-8 text-center">
                            <div class="text-2xl text-cyan-400/45">◇</div>
                            <p class="mt-2 text-xs font-semibold text-slate-500">Nenhuma fase cadastrada</p>
                            <button id="phaseEmptyAddBtn" type="button" class="mt-3 rounded-lg border border-cyan-400/20 px-3 py-2 text-[10px] font-bold text-cyan-300">Criar primeira fase</button>
                        </div>
                    </section>

                    <div class="grid gap-5 2xl:grid-cols-[minmax(0,1.2fr)_minmax(360px,.8fr)]">
                        <!-- DROP ZONE -->
                        <section class="panel rounded-2xl p-5">
                            <div class="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                                <div>
                                    <h3 class="text-sm font-bold text-slate-200">Ferramentas usadas</h3>
                                    <p class="mt-0.5 text-[10px] text-slate-600">Apps + Usados, Apps Locais e Links associados ao projeto</p>
                                </div>
                                <div class="flex items-center gap-2">
                                    <button
                                        id="importProjectToolsBtn"
                                        type="button"
                                        class="h-8 rounded-lg border border-cyan-400/15 bg-cyan-400/5 px-3 text-[10px] font-bold text-cyan-200 hover:bg-cyan-400/10 disabled:cursor-not-allowed disabled:opacity-40"
                                    >
                                        ⇩ Trazer de outro projeto
                                    </button>
                                    <span id="projectResourcesCount" class="rounded-lg border border-slate-700/25 bg-[#061722] px-2 py-1 text-[10px] text-slate-500">0</span>
                                </div>
                            </div>

                            <div id="projectDropZone" class="project-drop-zone rounded-2xl p-4">
                                <div id="projectResources" class="grid grid-cols-1 gap-2 sm:grid-cols-2 xl:grid-cols-3"></div>

                                <div id="projectDropEmpty" class="flex min-h-[140px] flex-col items-center justify-center text-center">
                                    <div class="text-3xl text-cyan-400/50">⇣</div>
                                    <p class="mt-2 text-xs font-semibold text-slate-400">Arraste ferramentas ou links para cá</p>
                                    <p class="mt-1 text-[10px] text-slate-600">A associação será salva automaticamente.</p>
                                </div>
                            </div>
                        </section>

                        <!-- BIBLIOTECA -->
                        <section class="panel rounded-2xl p-5">
                            <div class="mb-4 flex items-center justify-between gap-3">
                                <div>
                                    <h3 class="text-sm font-bold text-slate-200">Biblioteca disponível</h3>
                                    <p class="mt-0.5 text-[10px] text-slate-600">Arraste para o projeto</p>
                                </div>
                            </div>

                            <div class="grid grid-cols-3 gap-2">
                                <button data-library-tab="app" class="library-tab min-h-9 rounded-lg border border-cyan-400/25 bg-cyan-400/10 px-2 py-2 text-[10px] font-bold text-cyan-200 sm:text-xs">Apps + Usados</button>
                                <button data-library-tab="local_app" class="library-tab min-h-9 rounded-lg border border-slate-700/25 bg-[#071722] px-2 py-2 text-[10px] font-bold text-slate-500 sm:text-xs">Apps Locais</button>
                                <button data-library-tab="link" class="library-tab min-h-9 rounded-lg border border-slate-700/25 bg-[#071722] px-2 py-2 text-[10px] font-bold text-slate-500 sm:text-xs">Links</button>
                            </div>

                            <div class="mt-3 grid gap-2 sm:grid-cols-[minmax(0,1fr)_180px]">
                                <input id="resourceSearch" type="search" placeholder="Buscar recurso..." class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                                <select id="resourceCategoryFilter" class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-300 outline-none focus:border-cyan-400/35">
                                    <option value="all">Todas as categorias</option>
                                </select>
                            </div>

                            <div id="resourceLibrary" class="mt-3 max-h-[520px] space-y-2 overflow-y-auto pr-1"></div>
                            <div id="resourceLibraryEmpty" class="hidden py-10 text-center text-xs text-slate-600">Nenhum recurso encontrado.</div>
                        </section>
                    </div>
                </section>

                <section id="projectWelcome" class="panel rounded-2xl px-6 py-20 text-center">
                    <div class="text-5xl">◈</div>
                    <h3 class="mt-4 text-lg font-bold text-slate-200">Selecione uma pasta cadastrada</h3>
                    <p class="mx-auto mt-2 max-w-xl text-sm leading-6 text-slate-600">Clique em uma pasta com status para abrir o dashboard. Pastas “Sem cadastro” podem ser registradas sem mover ou alterar arquivos.</p>
                </section>
            </div>
        </main>
    </div>

    <!-- MODAL PROJETO -->
    <div id="projectModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-2xl rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 id="projectModalTitle" class="text-xl font-bold text-slate-100">Novo projeto</h3>
                    <p class="mt-1 text-xs text-slate-500">Defina o cliente, período e status do projeto.</p>
                </div>
                <button id="closeProjectModalBtn" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <form id="projectForm" class="space-y-4">
                <div class="grid gap-4 sm:grid-cols-2">
                    <div>
                        <label class="mb-1.5 block text-xs font-semibold text-slate-400">Projeto</label>
                        <input id="projectNameInput" required maxlength="120" placeholder="Projeto X" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                    </div>
                    <div>
                        <label class="mb-1.5 block text-xs font-semibold text-slate-400">Cliente</label>
                        <input id="projectClientInput" maxlength="120" placeholder="Nome do cliente" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                    </div>
                </div>

                <div class="grid gap-4 sm:grid-cols-3">
                    <div>
                        <label class="mb-1.5 block text-xs font-semibold text-slate-400">Data início</label>
                        <input id="projectStartInput" type="date" required class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                    </div>
                    <div>
                        <label class="mb-1.5 block text-xs font-semibold text-slate-400">Data final</label>
                        <input id="projectEndInput" type="date" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                    </div>
                    <div>
                        <label class="mb-1.5 block text-xs font-semibold text-slate-400">Status</label>
                        <select id="projectStatusInput" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                            <option value="planejado">Planejado</option>
                            <option value="em_andamento" selected>Em andamento</option>
                            <option value="pausado">Pausado</option>
                            <option value="concluido">Concluído</option>
                            <option value="cancelado">Cancelado</option>
                        </select>
                    </div>
                </div>

                <div>
                    <label class="mb-1.5 block text-xs font-semibold text-slate-400">Observações <span class="font-normal text-slate-600">(opcional)</span></label>
                    <textarea id="projectNotesInput" maxlength="500" rows="3" placeholder="Contexto, objetivo, entregas..." class="w-full resize-none rounded-xl border border-slate-700/30 bg-[#06131d] px-3 py-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35"></textarea>
                </div>

                <div id="projectFormError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:justify-end">
                    <button id="cancelProjectBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 hover:bg-slate-800/35 hover:text-slate-200">Cancelar</button>
                    <button id="saveProjectBtn" type="submit" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15 disabled:opacity-50">Salvar projeto</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL - TRAZER FERRAMENTAS DE OUTRO PROJETO -->
    <div id="importProjectToolsModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-md rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 class="text-lg font-bold text-slate-100">Trazer ferramentas de outro projeto</h3>
                    <p class="mt-1 text-xs text-slate-500">Escolha o projeto de origem. Recursos já existentes serão ignorados.</p>
                </div>
                <button id="closeImportProjectToolsBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <div>
                <label class="mb-1.5 block text-xs font-semibold text-slate-400">Projeto de origem</label>
                <select id="sourceProjectSelect" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35"></select>
            </div>

            <div id="sourceProjectSummary" class="mt-3 rounded-xl border border-slate-700/20 bg-[#061722] px-4 py-3 text-xs text-slate-500">
                Selecione um projeto.
            </div>

            <div id="importProjectToolsError" class="mt-4 hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

            <div class="mt-5 flex flex-col-reverse gap-2 sm:flex-row sm:justify-end">
                <button id="cancelImportProjectToolsBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400">Cancelar</button>
                <button id="runImportProjectToolsBtn" type="button" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 disabled:cursor-not-allowed disabled:opacity-40">
                    Trazer ferramentas
                </button>
            </div>
        </div>
    </div>

    <!-- MODAL FASE -->
    <div id="phaseModal" class="modal-backdrop fixed inset-0 z-50 hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="w-full max-w-md rounded-2xl border border-slate-700/30 bg-[#091d2a] p-6 shadow-2xl">
            <div class="mb-5 flex items-start justify-between gap-4">
                <div>
                    <h3 id="phaseModalTitle" class="text-lg font-bold text-slate-100">Nova fase</h3>
                    <p class="mt-1 text-xs text-slate-500">Defina a etapa e seu percentual de andamento.</p>
                </div>
                <button id="closePhaseModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <form id="phaseForm" class="space-y-4">
                <div>
                    <label class="mb-1.5 block text-xs font-semibold text-slate-400">Nome da fase</label>
                    <input id="phaseNameInput" required maxlength="100" placeholder="Ex.: Desenvolvimento" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-200 outline-none placeholder:text-slate-700 focus:border-cyan-400/35">
                </div>

                <div>
                    <label class="mb-1.5 block text-xs font-semibold text-slate-400">Status</label>
                    <select id="phaseStatusInput" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#06131d] px-3 text-sm text-slate-300 outline-none focus:border-cyan-400/35">
                        <option value="planejado">Planejado</option>
                        <option value="em_andamento">Em andamento</option>
                        <option value="pausado">Pausado</option>
                        <option value="concluido">Concluído</option>
                    </select>
                </div>

                <div>
                    <div class="mb-1.5 flex items-center justify-between gap-3">
                        <label class="text-xs font-semibold text-slate-400">Andamento</label>
                        <span id="phaseProgressValue" class="text-xs font-bold text-cyan-300">0%</span>
                    </div>
                    <input id="phaseProgressInput" type="range" min="0" max="100" step="5" value="0" class="w-full accent-cyan-400">
                </div>

                <div id="phaseFormError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 pt-1 sm:flex-row sm:justify-end">
                    <button id="cancelPhaseBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400">Cancelar</button>
                    <button id="savePhaseBtn" type="submit" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200">Salvar fase</button>
                </div>
            </form>
        </div>
    </div>

    <div id="workspaceToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const STATUS_LABELS = {
            planejado: 'Planejado',
            em_andamento: 'Em andamento',
            pausado: 'Pausado',
            concluido: 'Concluído',
            cancelado: 'Cancelado'
        };

        const state = {
            projects: [],
            apps: [],
            localApps: [],
            links: [],
            workspaceFolders: [],
            workspaceSummary: {},
            workspaceRoot: '',
            workspaceFilter: 'all',
            workspaceSearch: '',
            selectedFolderName: '',
            selectedProjectId: null,
            libraryTab: 'app',
            editProjectId: null,
            editPhaseId: null
        };

        const projectsList = document.getElementById('projectsList');
        const projectsEmpty = document.getElementById('projectsEmpty');
        const projectsCount = document.getElementById('projectsCount');
        const workspaceExplorerList = document.getElementById('workspaceExplorerList');
        const workspaceExplorerEmpty = document.getElementById('workspaceExplorerEmpty');
        const workspaceExplorerRoot = document.getElementById('workspaceExplorerRoot');
        const workspaceExplorerSearch = document.getElementById('workspaceExplorerSearch');
        const workspaceExplorerCount = document.getElementById('workspaceExplorerCount');
        const workspaceExplorerSelection = document.getElementById('workspaceExplorerSelection');
        const projectArea = document.getElementById('projectArea');
        const projectWelcome = document.getElementById('projectWelcome');
        const projectTitle = document.getElementById('projectTitle');
        const projectClient = document.getElementById('projectClient');
        const projectStart = document.getElementById('projectStart');
        const projectEnd = document.getElementById('projectEnd');
        const projectStatusBadge = document.getElementById('projectStatusBadge');
        const projectNotes = document.getElementById('projectNotes');
        const projectDropZone = document.getElementById('projectDropZone');
        const projectResources = document.getElementById('projectResources');
        const projectDropEmpty = document.getElementById('projectDropEmpty');
        const projectResourcesCount = document.getElementById('projectResourcesCount');

        const dashboardProgress = document.getElementById('dashboardProgress');
        const dashboardPhases = document.getElementById('dashboardPhases');
        const dashboardDeadline = document.getElementById('dashboardDeadline');
        const dashboardResources = document.getElementById('dashboardResources');
        const dashboardProgressLabel = document.getElementById('dashboardProgressLabel');
        const dashboardProgressBar = document.getElementById('dashboardProgressBar');
        const phaseCountBadge = document.getElementById('phaseCountBadge');
        const phaseTimeline = document.getElementById('phaseTimeline');
        const phaseEmpty = document.getElementById('phaseEmpty');

        const resourceLibrary = document.getElementById('resourceLibrary');
        const resourceLibraryEmpty = document.getElementById('resourceLibraryEmpty');
        const resourceSearch = document.getElementById('resourceSearch');
        const resourceCategoryFilter = document.getElementById('resourceCategoryFilter');
        const importProjectToolsBtn = document.getElementById('importProjectToolsBtn');
        const workspaceToast = document.getElementById('workspaceToast');

        const importProjectToolsModal = document.getElementById('importProjectToolsModal');
        const sourceProjectSelect = document.getElementById('sourceProjectSelect');
        const sourceProjectSummary = document.getElementById('sourceProjectSummary');
        const importProjectToolsError = document.getElementById('importProjectToolsError');
        const runImportProjectToolsBtn = document.getElementById('runImportProjectToolsBtn');

        const projectModal = document.getElementById('projectModal');
        const projectForm = document.getElementById('projectForm');
        const projectModalTitle = document.getElementById('projectModalTitle');
        const projectNameInput = document.getElementById('projectNameInput');
        const projectClientInput = document.getElementById('projectClientInput');
        const projectStartInput = document.getElementById('projectStartInput');
        const projectEndInput = document.getElementById('projectEndInput');
        const projectStatusInput = document.getElementById('projectStatusInput');
        const projectNotesInput = document.getElementById('projectNotesInput');
        const projectFormError = document.getElementById('projectFormError');
        const saveProjectBtn = document.getElementById('saveProjectBtn');

        const phaseModal = document.getElementById('phaseModal');
        const phaseForm = document.getElementById('phaseForm');
        const phaseModalTitle = document.getElementById('phaseModalTitle');
        const phaseNameInput = document.getElementById('phaseNameInput');
        const phaseStatusInput = document.getElementById('phaseStatusInput');
        const phaseProgressInput = document.getElementById('phaseProgressInput');
        const phaseProgressValue = document.getElementById('phaseProgressValue');
        const phaseFormError = document.getElementById('phaseFormError');
        const savePhaseBtn = document.getElementById('savePhaseBtn');

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function normalize(value) {
            return String(value ?? '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase();
        }

        function safeUrl(value) {
            try {
                const url = new URL(value);
                return ['http:', 'https:'].includes(url.protocol) ? url.href : '#';
            } catch {
                return '#';
            }
        }

        function iconMarkup(icon, name, sizeClass = 'h-8 w-8') {
            const value = String(icon || '').trim();
            if (/^(https?:\/\/|\/icons\/)/i.test(value)) {
                return `<img src="${escapeHtml(value)}" alt="${escapeHtml(name)}" class="${sizeClass} object-contain">`;
            }
            return `<span class="text-2xl">${escapeHtml(value || '◈')}</span>`;
        }

        function linkFavicon(url) {
            const safe = safeUrl(url);
            if (safe === '#') return '';
            return `https://www.google.com/s2/favicons?domain_url=${encodeURIComponent(safe)}&sz=64`;
        }

        function showToast(message, error = false) {
            workspaceToast.textContent = message;
            workspaceToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            workspaceToast.classList.remove('hidden');
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => workspaceToast.classList.add('hidden'), 2300);
        }

        function selectedProject() {
            return state.projects.find(project => Number(project.id) === Number(state.selectedProjectId)) || null;
        }

        function formatDate(value) {
            if (!value) return '—';
            const parts = String(value).split('-');
            return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : value;
        }

        function projectProgress(project) {
            const phases = Array.isArray(project?.phases) ? project.phases : [];
            if (!phases.length) return 0;

            const total = phases.reduce((sum, phase) => {
                const progress = Math.max(0, Math.min(100, Number(phase.progress) || 0));
                return sum + progress;
            }, 0);

            return Math.round(total / phases.length);
        }

        function deadlineSummary(endDate) {
            if (!endDate) return 'Sem prazo';

            const parts = String(endDate).split('-').map(Number);
            if (parts.length !== 3 || parts.some(Number.isNaN)) return '—';

            const target = Date.UTC(parts[0], parts[1] - 1, parts[2]);
            const now = new Date();
            const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
            const days = Math.round((target - today) / 86400000);

            if (days === 0) return 'Hoje';
            if (days > 0) return `${days} dia${days === 1 ? '' : 's'}`;
            const overdue = Math.abs(days);
            return `${overdue} dia${overdue === 1 ? '' : 's'} atrasado${overdue === 1 ? '' : 's'}`;
        }

        function renderProjectDashboard(project, resourceCount) {
            const phases = Array.isArray(project.phases) ? project.phases : [];
            const progress = projectProgress(project);
            const completed = phases.filter(phase =>
                phase.status === 'concluido' || Number(phase.progress) >= 100
            ).length;

            dashboardProgress.textContent = `${progress}%`;
            dashboardPhases.textContent = `${completed}/${phases.length}`;
            dashboardDeadline.textContent = deadlineSummary(project.end_date);
            dashboardResources.textContent = resourceCount;
            dashboardProgressLabel.textContent = `${progress}%`;
            dashboardProgressBar.style.width = `${progress}%`;
            phaseCountBadge.textContent = `${phases.length} ${phases.length === 1 ? 'fase' : 'fases'}`;

            phaseEmpty.classList.toggle('hidden', phases.length > 0);

            phaseTimeline.innerHTML = phases.map((phase, index) => {
                const phaseProgress = Math.max(0, Math.min(100, Number(phase.progress) || 0));
                return `
                    <article class="phase-card rounded-xl p-3">
                        <div class="flex items-start justify-between gap-2">
                            <div class="min-w-0">
                                <div class="text-[9px] font-bold uppercase tracking-[.12em] text-cyan-400/55">Fase ${index + 1}</div>
                                <div class="mt-1 truncate text-xs font-bold text-slate-300">${escapeHtml(phase.name)}</div>
                            </div>
                            <div class="flex shrink-0 items-center gap-1">
                                <button type="button" data-edit-phase="${Number(phase.id)}" class="rounded-md p-1.5 text-[11px] text-slate-600 hover:bg-cyan-950/30 hover:text-cyan-300" title="Editar fase">✎</button>
                                <button type="button" data-delete-phase="${Number(phase.id)}" class="rounded-md p-1.5 text-[11px] text-slate-700 hover:bg-red-950/30 hover:text-red-300" title="Excluir fase">✕</button>
                            </div>
                        </div>

                        <div class="mt-3 flex items-center justify-between gap-2">
                            <span class="rounded-md border px-2 py-1 text-[9px] font-bold uppercase tracking-[.05em] status-${escapeHtml(phase.status)}">
                                ${escapeHtml(STATUS_LABELS[phase.status] || phase.status)}
                            </span>
                            <span class="text-[10px] font-bold text-cyan-300">${phaseProgress}%</span>
                        </div>

                        <div class="phase-progress-track mt-3">
                            <div class="phase-progress-fill" style="width:${phaseProgress}%"></div>
                        </div>
                    </article>
                `;
            }).join('');
        }

        function resourceObject(resource) {
            if (resource.resource_type === 'app') {
                const app = state.apps.find(item => Number(item.id) === Number(resource.resource_id));
                return app ? {
                    type: 'app',
                    id: Number(app.id),
                    title: app.name,
                    subtitle: app.category || 'App',
                    category: app.category || 'Outros',
                    url: app.url,
                    icon: app.icon
                } : null;
            }

            if (resource.resource_type === 'local_app') {
                const localApp = state.localApps.find(
                    item => String(item.id) === String(resource.resource_id)
                );

                return localApp ? {
                    type: 'local_app',
                    id: String(localApp.id),
                    title: localApp.name,
                    subtitle: localApp.group || 'App local',
                    category: localApp.group || 'Outros',
                    url: '',
                    icon: localApp.icon,
                    installed: Boolean(localApp.installed)
                } : null;
            }

            const link = state.links.find(item => Number(item.id) === Number(resource.resource_id));
            return link ? {
                type: 'link',
                id: Number(link.id),
                title: link.title,
                subtitle: link.category || 'Link',
                category: link.category || 'Outros',
                url: link.url,
                icon: linkFavicon(link.url)
            } : null;
        }

        function workspaceStatusClass(status) {
            if (status === 'concluido') return 'status-concluido';
            if (status === 'pausado') return 'status-pausado';
            if (status === 'cancelado') return 'status-cancelado';
            if (status === 'planejado') return 'status-planejado';
            if (status === 'em_andamento') return 'status-em_andamento';
            return 'border-slate-700/30 bg-slate-800/20 text-slate-500';
        }

        function formatWorkspaceModified(timestamp) {
            const numeric = Number(timestamp || 0);
            if (!numeric) return '—';
            const date = new Date(numeric * 1000);
            if (Number.isNaN(date.getTime())) return '—';
            return date.toLocaleString('pt-BR', {
                day: '2-digit', month: '2-digit', year: 'numeric',
                hour: '2-digit', minute: '2-digit'
            });
        }

        function projectForFolder(folder) {
            if (!folder?.project_id) return null;
            return state.projects.find(
                project => Number(project.id) === Number(folder.project_id)
            ) || null;
        }

        function filteredWorkspaceFolders() {
            const query = normalize(state.workspaceSearch || '');
            return state.workspaceFolders.filter(folder => {
                const filterMatch = state.workspaceFilter === 'all'
                    || (state.workspaceFilter === 'git'
                        ? Boolean(folder.is_git)
                        : folder.status === state.workspaceFilter);
                if (!filterMatch) return false;
                if (!query) return true;
                return normalize([
                    folder.name,
                    folder.client,
                    STATUS_LABELS[folder.status] || folder.status
                ].join(' ')).includes(query);
            });
        }

        function renderWorkspaceExplorer() {
            const folders = filteredWorkspaceFolders();
            const summary = state.workspaceSummary || {};

            workspaceExplorerRoot.textContent = state.workspaceRoot || 'D:\\python\\CHATGPT';
            workspaceExplorerCount.textContent = `${folders.length} exibido(s) • ${state.workspaceFolders.length} total`;
            workspaceExplorerEmpty.classList.toggle('hidden', folders.length > 0);

            document.getElementById('workspaceCountAll').textContent = summary.total ?? 0;
            document.getElementById('workspaceCountActive').textContent = summary.em_andamento ?? 0;
            document.getElementById('workspaceCountPaused').textContent = summary.pausado ?? 0;
            document.getElementById('workspaceCountDone').textContent = summary.concluido ?? 0;
            document.getElementById('workspaceCountUnregistered').textContent = summary.nao_classificado ?? 0;
            document.getElementById('workspaceCountGit').textContent = summary.git ?? 0;

            document.querySelectorAll('[data-workspace-filter]').forEach(button => {
                button.classList.toggle(
                    'active',
                    button.dataset.workspaceFilter === state.workspaceFilter
                );
            });

            workspaceExplorerList.innerHTML = folders.map(folder => {
                const selected = folder.name === state.selectedFolderName;
                const statusLabel = folder.registered
                    ? (STATUS_LABELS[folder.status] || folder.status)
                    : 'Sem cadastro';
                const progress = Number(folder.progress || 0);

                return `
                    <div class="explorer-row ${selected ? 'selected' : ''} grid cursor-default grid-cols-[minmax(260px,1.7fr)_150px_110px_90px_150px_130px] items-center px-4 py-2.5" data-workspace-folder="${escapeHtml(folder.name)}" data-project-id="${folder.project_id ? Number(folder.project_id) : ''}">
                        <div class="flex min-w-0 items-center gap-3 pr-3">
                            <div class="folder-icon shrink-0"></div>
                            <div class="min-w-0">
                                <div class="truncate text-xs font-semibold text-slate-300">${escapeHtml(folder.name)}</div>
                                <div class="mt-0.5 truncate text-[9px] text-slate-700">${folder.item_count} ${folder.item_count === 1 ? 'item' : 'itens'}${folder.client ? ` • ${escapeHtml(folder.client)}` : ''}</div>
                            </div>
                        </div>
                        <div class="explorer-hide-mobile pr-3"><span class="inline-flex rounded-lg border px-2 py-1 text-[9px] font-bold uppercase tracking-[.06em] ${workspaceStatusClass(folder.status)}">${escapeHtml(statusLabel)}</span></div>
                        <div class="explorer-hide-mobile pr-3">${folder.registered ? `<div class="text-[10px] font-bold text-cyan-300/80">${progress}%</div><div class="mt-1 h-1 overflow-hidden rounded-full bg-[#061722]"><div class="overall-progress-fill" style="width:${progress}%"></div></div>` : '<span class="text-[10px] text-slate-700">—</span>'}</div>
                        <div class="explorer-hide-mobile">${folder.is_git ? '<span class="rounded-md border border-violet-400/20 bg-violet-400/5 px-2 py-1 text-[9px] font-bold text-violet-200">⑂ Git</span>' : '<span class="text-[10px] text-slate-700">—</span>'}</div>
                        <div class="explorer-hide-mobile text-[10px] text-slate-500">${escapeHtml(formatWorkspaceModified(folder.modified_ts))}</div>
                        <div class="flex items-center justify-end gap-1">
                            <button data-workspace-action="open" data-folder="${escapeHtml(folder.name)}" class="grid h-7 w-7 place-items-center rounded-md text-[11px] text-slate-500 hover:bg-cyan-400/5 hover:text-cyan-200" title="Abrir no Explorer">📂</button>
                            <button data-workspace-action="terminal" data-folder="${escapeHtml(folder.name)}" class="grid h-7 w-7 place-items-center rounded-md text-[11px] text-slate-500 hover:bg-cyan-400/5 hover:text-cyan-200" title="Terminal">⌘</button>
                            <button data-workspace-action="git" data-folder="${escapeHtml(folder.name)}" class="grid h-7 w-7 place-items-center rounded-md text-[11px] text-slate-500 hover:bg-violet-400/5 hover:text-violet-200" title="Git & GitHub">⑂</button>
                            ${folder.registered
                                ? `<button data-workspace-action="details" data-folder="${escapeHtml(folder.name)}" data-project-id="${Number(folder.project_id)}" class="grid h-7 w-7 place-items-center rounded-md text-[11px] text-slate-500 hover:bg-emerald-400/5 hover:text-emerald-200" title="Detalhes">◈</button>`
                                : `<button data-workspace-action="register" data-folder="${escapeHtml(folder.name)}" class="grid h-7 w-7 place-items-center rounded-md text-[11px] text-slate-500 hover:bg-emerald-400/5 hover:text-emerald-200" title="Cadastrar">＋</button>`
                            }
                        </div>
                        <div class="mt-3 flex flex-wrap items-center gap-2 md:hidden"><span class="rounded-lg border px-2 py-1 text-[9px] font-bold uppercase tracking-[.06em] ${workspaceStatusClass(folder.status)}">${escapeHtml(statusLabel)}</span>${folder.is_git ? '<span class="text-[9px] font-bold text-violet-300">⑂ Git</span>' : ''}${folder.registered ? `<span class="text-[9px] font-bold text-cyan-300">${progress}%</span>` : ''}<span class="text-[9px] text-slate-600">${escapeHtml(formatWorkspaceModified(folder.modified_ts))}</span></div>
                    </div>
                `;
            }).join('');
        }

        async function loadWorkspaceExplorer() {
            const response = await fetch('/api/workspace/explorer', { cache: 'no-store' });
            const data = await response.json();
            if (!response.ok) {
                throw new Error(data.detail || 'Não foi possível ler as pastas do Workspace.');
            }
            state.workspaceFolders = Array.isArray(data.items) ? data.items : [];
            state.workspaceSummary = data.summary || {};
            state.workspaceRoot = data.root || 'D:\\python\\CHATGPT';
            if (!state.workspaceFolders.some(folder => folder.name === state.selectedFolderName)) {
                state.selectedFolderName = '';
            }
            renderWorkspaceExplorer();
        }

        async function runWorkspaceFolderAction(folderName, action) {
            try {
                const response = await fetch('/api/workspace/explorer/action', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ folder_name: folderName, action })
                });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Ação não concluída.');
                showToast(data.message || 'Ação concluída.');
                if (action === 'git' && data.navigate_to) window.location.href = data.navigate_to;
            } catch (error) {
                showToast(error.message || 'Erro na ação da pasta.', true);
            }
        }

        function renderProjects() {
            projectsCount.textContent = `${state.projects.length} ${state.projects.length === 1 ? 'projeto' : 'projetos'}`;
            projectsEmpty.classList.toggle('hidden', state.projects.length > 0);

            projectsList.innerHTML = state.projects.map(project => {
                const active = Number(project.id) === Number(state.selectedProjectId);
                return `
                    <button
                        type="button"
                        data-project-id="${Number(project.id)}"
                        class="project-card ${active ? 'active' : ''} w-full rounded-xl p-3 text-left"
                    >
                        <div class="flex items-start justify-between gap-2">
                            <div class="min-w-0">
                                <div class="truncate text-xs font-bold text-slate-300">${escapeHtml(project.name)}</div>
                                <div class="mt-1 truncate text-[10px] text-slate-600">${escapeHtml(project.client || 'Sem cliente')}</div>
                            </div>
                            <span class="h-2 w-2 shrink-0 rounded-full ${
                                project.status === 'concluido' ? 'bg-emerald-400' :
                                project.status === 'pausado' ? 'bg-amber-300' :
                                project.status === 'cancelado' ? 'bg-red-400' :
                                project.status === 'planejado' ? 'bg-blue-400' :
                                'bg-cyan-400'
                            }"></span>
                        </div>
                        <div class="mt-2 flex items-center justify-between gap-2">
                            <span class="text-[9px] uppercase tracking-[.06em] text-slate-600">${escapeHtml(STATUS_LABELS[project.status] || project.status)}</span>
                            <span class="text-[9px] font-bold text-cyan-400/60">${projectProgress(project)}%</span>
                        </div>
                        <div class="mt-2 h-1 overflow-hidden rounded-full bg-[#061722]">
                            <div class="overall-progress-fill" style="width:${projectProgress(project)}%"></div>
                        </div>
                    </button>
                `;
            }).join('');
        }

        function renderSelectedProject() {
            const project = selectedProject();

            if (!project) {
                projectArea.classList.add('hidden');
                projectWelcome.classList.remove('hidden');
                return;
            }

            projectArea.classList.remove('hidden');
            projectWelcome.classList.add('hidden');

            projectTitle.textContent = project.name;
            projectClient.textContent = project.client || '—';
            projectStart.textContent = formatDate(project.start_date);
            projectEnd.textContent = formatDate(project.end_date);
            projectStatusBadge.textContent = STATUS_LABELS[project.status] || project.status;
            projectStatusBadge.className = `rounded-lg border px-2.5 py-1 text-[10px] font-bold uppercase tracking-[.08em] status-${project.status}`;

            projectNotes.textContent = project.notes || '';
            projectNotes.classList.toggle('hidden', !project.notes);

            const resources = (project.resources || [])
                .map(resourceObject)
                .filter(Boolean);

            projectResourcesCount.textContent = resources.length;
            projectDropEmpty.classList.toggle('hidden', resources.length > 0);
            importProjectToolsBtn.disabled = state.projects.filter(
                item => Number(item.id) !== Number(project.id)
            ).length === 0;
            renderProjectDashboard(project, resources.length);

            projectResources.innerHTML = resources.map(resource => `
                <article class="resource-card group flex items-center gap-3 rounded-xl p-3">
                    <div class="grid h-10 w-10 shrink-0 place-items-center overflow-hidden rounded-lg border border-slate-700/25 bg-[#061722]">
                        ${
                            resource.type === 'link'
                                ? (resource.icon
                                    ? `<img src="${escapeHtml(resource.icon)}" alt="" class="h-7 w-7 object-contain">`
                                    : '<span>🔗</span>')
                                : iconMarkup(resource.icon, resource.title)
                        }
                    </div>

                    <div class="min-w-0 flex-1">
                        <div class="truncate text-xs font-semibold text-slate-300">${escapeHtml(resource.title)}</div>
                        <div class="mt-0.5 truncate text-[10px] text-slate-600">${escapeHtml(resource.subtitle)}</div>
                    </div>

                    ${
                        resource.type === 'local_app'
                            ? `<button
                                    type="button"
                                    data-launch-local-id="${escapeHtml(resource.id)}"
                                    ${resource.installed ? '' : 'disabled'}
                                    class="rounded-lg border px-2 py-1.5 text-[10px] font-bold ${
                                        resource.installed
                                            ? 'border-cyan-400/15 bg-cyan-400/5 text-cyan-300 hover:bg-cyan-400/10'
                                            : 'cursor-not-allowed border-slate-800/30 text-slate-700'
                                    }"
                               >${resource.installed ? 'Abrir' : 'Indisponível'}</button>`
                            : `<a
                                    href="${escapeHtml(safeUrl(resource.url))}"
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    class="rounded-lg px-2 py-1.5 text-[10px] font-bold text-cyan-400/65 hover:bg-cyan-400/5 hover:text-cyan-200"
                               >Abrir ↗</a>`
                    }

                    <button
                        type="button"
                        data-remove-resource-type="${resource.type}"
                        data-remove-resource-id="${escapeHtml(resource.id)}"
                        class="rounded-lg p-2 text-[11px] text-slate-700 opacity-70 hover:bg-red-950/30 hover:text-red-300"
                        title="Remover do projeto"
                    >✕</button>
                </article>
            `).join('');

            renderLibrary();
        }

        function assignedResourceKeys() {
            const project = selectedProject();
            return new Set(
                (project?.resources || []).map(resource =>
                    `${resource.resource_type}:${String(resource.resource_id)}`
                )
            );
        }

        function libraryResourcesForCurrentTab() {
            if (state.libraryTab === 'app') {
                return [...state.apps]
                    .sort((a, b) => Number(Boolean(b.favorite)) - Number(Boolean(a.favorite)))
                    .map(app => ({
                        type: 'app',
                        id: Number(app.id),
                        title: app.name,
                        subtitle: app.category || 'App',
                        category: app.category || 'Outros',
                        url: app.url,
                        icon: app.icon
                    }));
            }

            if (state.libraryTab === 'local_app') {
                return state.localApps
                    .filter(app => Boolean(app.installed))
                    .map(app => ({
                        type: 'local_app',
                        id: String(app.id),
                        title: app.name,
                        subtitle: app.group || 'App local',
                        category: app.group || 'Outros',
                        url: '',
                        icon: app.icon
                    }));
            }

            return state.links.map(link => ({
                type: 'link',
                id: Number(link.id),
                title: link.title,
                subtitle: link.category || 'Link',
                category: link.category || 'Outros',
                url: link.url,
                icon: linkFavicon(link.url)
            }));
        }

        function populateResourceCategoryFilter(reset = false) {
            const current = reset ? 'all' : (resourceCategoryFilter.value || 'all');

            const categories = [...new Set(
                libraryResourcesForCurrentTab()
                    .map(resource => String(resource.category || 'Outros').trim())
                    .filter(Boolean)
            )].sort((a, b) => a.localeCompare(b, 'pt-BR', { sensitivity: 'base' }));

            resourceCategoryFilter.innerHTML = [
                '<option value="all">Todas as categorias</option>',
                ...categories.map(category =>
                    `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`
                )
            ].join('');

            const exists = categories.some(category => category === current);
            resourceCategoryFilter.value = exists ? current : 'all';
        }

        function renderLibrary() {
            if (!selectedProject()) return;

            const query = normalize(resourceSearch.value.trim());
            const selectedCategory = resourceCategoryFilter.value || 'all';
            const assigned = assignedResourceKeys();

            let resources = libraryResourcesForCurrentTab();

            resources = resources.filter(resource => {
                if (assigned.has(`${resource.type}:${String(resource.id)}`)) return false;

                const matchesQuery = !query ||
                    [resource.title, resource.subtitle, resource.category]
                        .some(value => normalize(value).includes(query));

                const matchesCategory = selectedCategory === 'all' ||
                    String(resource.category || 'Outros') === selectedCategory;

                return matchesQuery && matchesCategory;
            });

            resourceLibraryEmpty.classList.toggle('hidden', resources.length > 0);

            resourceLibrary.innerHTML = resources.map(resource => `
                <article
                    class="resource-card flex items-center gap-3 rounded-xl p-3"
                    draggable="true"
                    data-library-resource-type="${resource.type}"
                    data-library-resource-id="${escapeHtml(resource.id)}"
                >
                    <div class="grid h-10 w-10 shrink-0 place-items-center overflow-hidden rounded-lg border border-slate-700/25 bg-[#061722]">
                        ${
                            resource.type === 'link'
                                ? (resource.icon
                                    ? `<img src="${escapeHtml(resource.icon)}" alt="" class="h-7 w-7 object-contain">`
                                    : '<span>🔗</span>')
                                : iconMarkup(resource.icon, resource.title)
                        }
                    </div>

                    <div class="min-w-0 flex-1">
                        <div class="truncate text-xs font-semibold text-slate-300">${escapeHtml(resource.title)}</div>
                        <div class="mt-0.5 truncate text-[10px] text-slate-600">${escapeHtml(resource.subtitle)}</div>
                    </div>

                    <button
                        type="button"
                        data-add-resource-type="${resource.type}"
                        data-add-resource-id="${escapeHtml(resource.id)}"
                        class="grid h-8 w-8 place-items-center rounded-lg border border-cyan-400/15 bg-cyan-400/5 text-cyan-300 hover:bg-cyan-400/10"
                        title="Adicionar ao projeto"
                    >＋</button>
                </article>
            `).join('');
        }

        function openPhaseModal(phase = null) {
            const project = selectedProject();
            if (!project) return;

            state.editPhaseId = phase ? Number(phase.id) : null;
            phaseForm.reset();
            phaseFormError.classList.add('hidden');

            phaseModalTitle.textContent = phase ? 'Editar fase' : 'Nova fase';
            savePhaseBtn.textContent = phase ? 'Salvar alterações' : 'Salvar fase';

            if (phase) {
                phaseNameInput.value = phase.name || '';
                phaseStatusInput.value = phase.status || 'planejado';
                phaseProgressInput.value = Math.max(0, Math.min(100, Number(phase.progress) || 0));
            } else {
                phaseStatusInput.value = 'planejado';
                phaseProgressInput.value = 0;
            }

            phaseProgressValue.textContent = `${phaseProgressInput.value}%`;
            phaseModal.classList.remove('hidden');
            phaseModal.classList.add('flex');
            setTimeout(() => phaseNameInput.focus(), 50);
        }

        function closePhaseModal() {
            phaseModal.classList.add('hidden');
            phaseModal.classList.remove('flex');
            phaseForm.reset();
            phaseFormError.classList.add('hidden');
            state.editPhaseId = null;
        }

        async function savePhase(event) {
            event.preventDefault();

            const project = selectedProject();
            if (!project) return;

            const editing = state.editPhaseId !== null;
            const payload = {
                name: phaseNameInput.value.trim(),
                status: phaseStatusInput.value,
                progress: Number(phaseProgressInput.value)
            };

            savePhaseBtn.disabled = true;
            savePhaseBtn.textContent = 'Salvando...';
            phaseFormError.classList.add('hidden');

            try {
                const response = await fetch(
                    editing
                        ? `/api/projects/${project.id}/phases/${state.editPhaseId}`
                        : `/api/projects/${project.id}/phases`,
                    {
                        method: editing ? 'PUT' : 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    }
                );

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível salvar a fase.');
                }

                const index = state.projects.findIndex(item => Number(item.id) === Number(data.id));
                if (index !== -1) state.projects[index] = data;

                closePhaseModal();
                renderProjects();
                renderWorkspaceExplorer();
                renderSelectedProject();
                showToast(editing ? 'Fase atualizada.' : 'Fase adicionada.');
            } catch (error) {
                phaseFormError.textContent = error.message || 'Erro ao salvar fase.';
                phaseFormError.classList.remove('hidden');
            } finally {
                savePhaseBtn.disabled = false;
                savePhaseBtn.textContent = state.editPhaseId ? 'Salvar alterações' : 'Salvar fase';
            }
        }

        async function deletePhase(phaseId) {
            const project = selectedProject();
            if (!project) return;

            const phase = (project.phases || []).find(item => Number(item.id) === Number(phaseId));
            if (!phase) return;

            if (!window.confirm(`Excluir a fase “${phase.name}”?`)) return;

            try {
                const response = await fetch(
                    `/api/projects/${project.id}/phases/${phaseId}`,
                    { method: 'DELETE' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível excluir a fase.');
                }

                const index = state.projects.findIndex(item => Number(item.id) === Number(data.id));
                if (index !== -1) state.projects[index] = data;

                renderProjects();
                renderSelectedProject();
                showToast('Fase excluída.');
            } catch (error) {
                showToast(error.message || 'Erro ao excluir fase.', true);
            }
        }

        function openProjectModal(project = null, suggestedFolderName = '') {
            state.editProjectId = project ? Number(project.id) : null;
            projectForm.reset();
            projectFormError.classList.add('hidden');

            projectModalTitle.textContent = project ? 'Editar projeto' : 'Novo projeto';
            saveProjectBtn.textContent = project ? 'Salvar alterações' : 'Salvar projeto';

            if (project) {
                projectNameInput.value = project.name || '';
                projectClientInput.value = project.client || '';
                projectStartInput.value = project.start_date || '';
                projectEndInput.value = project.end_date || '';
                projectStatusInput.value = project.status || 'em_andamento';
                projectNotesInput.value = project.notes || '';
            } else {
                projectNameInput.value = suggestedFolderName || '';
                projectStatusInput.value = 'em_andamento';
                projectStartInput.value = new Date().toISOString().slice(0, 10);
            }

            projectModal.classList.remove('hidden');
            projectModal.classList.add('flex');
            setTimeout(() => projectNameInput.focus(), 50);
        }

        function closeProjectModal() {
            projectModal.classList.add('hidden');
            projectModal.classList.remove('flex');
            projectForm.reset();
            projectFormError.classList.add('hidden');
            state.editProjectId = null;
        }

        async function loadWorkspace() {
            try {
                const [projectsResponse, appsResponse, localAppsResponse, linksResponse, explorerResponse] = await Promise.all([
                    fetch('/api/projects', { cache: 'no-store' }),
                    fetch('/api/apps', { cache: 'no-store' }),
                    fetch('/api/local-apps', { cache: 'no-store' }),
                    fetch('/api/links', { cache: 'no-store' }),
                    fetch('/api/workspace/explorer', { cache: 'no-store' })
                ]);

                if (
                    !projectsResponse.ok ||
                    !appsResponse.ok ||
                    !localAppsResponse.ok ||
                    !linksResponse.ok ||
                    !explorerResponse.ok
                ) {
                    throw new Error('Não foi possível carregar os dados da área personalizada.');
                }

                state.projects = await projectsResponse.json();
                state.apps = await appsResponse.json();
                state.localApps = await localAppsResponse.json();
                state.links = await linksResponse.json();
                const explorerData = await explorerResponse.json();
                state.workspaceFolders = Array.isArray(explorerData.items) ? explorerData.items : [];
                state.workspaceSummary = explorerData.summary || {};
                state.workspaceRoot = explorerData.root || 'D:\\python\\CHATGPT';
                populateResourceCategoryFilter(true);


                renderProjects();
                renderSelectedProject();
            } catch (error) {
                showToast(error.message || 'Erro ao carregar projetos.', true);
            }
        }

        async function saveProject(event) {
            event.preventDefault();

            const editing = state.editProjectId !== null;
            const payload = {
                name: projectNameInput.value.trim(),
                client: projectClientInput.value.trim(),
                start_date: projectStartInput.value,
                end_date: projectEndInput.value,
                status: projectStatusInput.value,
                notes: projectNotesInput.value.trim()
            };

            saveProjectBtn.disabled = true;
            saveProjectBtn.textContent = editing ? 'Salvando...' : 'Criando...';
            projectFormError.classList.add('hidden');

            try {
                const response = await fetch(
                    editing ? `/api/projects/${state.editProjectId}` : '/api/projects',
                    {
                        method: editing ? 'PUT' : 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(payload)
                    }
                );

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível salvar o projeto.');
                }

                if (editing) {
                    const index = state.projects.findIndex(project => Number(project.id) === Number(data.id));
                    if (index !== -1) state.projects[index] = data;
                } else {
                    state.projects.push(data);
                    state.selectedProjectId = Number(data.id);
                }

                closeProjectModal();
                await loadWorkspaceExplorer();
                renderProjects();
                renderSelectedProject();
                showToast(editing ? 'Projeto atualizado.' : 'Projeto criado.');
            } catch (error) {
                projectFormError.textContent = error.message || 'Erro ao salvar projeto.';
                projectFormError.classList.remove('hidden');
            } finally {
                saveProjectBtn.disabled = false;
                saveProjectBtn.textContent = state.editProjectId ? 'Salvar alterações' : 'Salvar projeto';
            }
        }

        async function deleteSelectedProject() {
            const project = selectedProject();
            if (!project) return;

            if (!window.confirm(`Excluir o projeto “${project.name}”?`)) return;

            try {
                const response = await fetch(`/api/projects/${project.id}`, {
                    method: 'DELETE'
                });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível excluir o projeto.');
                }

                state.projects = state.projects.filter(item => Number(item.id) !== Number(project.id));
                state.selectedProjectId = null;
                state.selectedFolderName = '';
                await loadWorkspaceExplorer();
                renderProjects();
                renderSelectedProject();
                showToast('Cadastro excluído. A pasta física foi preservada.');
            } catch (error) {
                showToast(error.message || 'Erro ao excluir projeto.', true);
            }
        }

        async function launchWorkspaceLocalApp(localAppId) {
            try {
                const response = await fetch(
                    `/api/local-apps/${encodeURIComponent(localAppId)}/launch`,
                    { method: 'POST' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível abrir o aplicativo local.');
                }

                showToast(data.message || 'Aplicativo aberto.');
            } catch (error) {
                showToast(error.message || 'Erro ao abrir aplicativo local.', true);
            }
        }

        function updateSourceProjectSummary() {
            const sourceId = Number(sourceProjectSelect.value);
            const source = state.projects.find(project => Number(project.id) === sourceId);

            if (!source) {
                sourceProjectSummary.textContent = 'Selecione um projeto.';
                runImportProjectToolsBtn.disabled = true;
                return;
            }

            const resources = Array.isArray(source.resources) ? source.resources : [];
            const apps = resources.filter(resource => resource.resource_type === 'app').length;
            const locals = resources.filter(resource => resource.resource_type === 'local_app').length;
            const links = resources.filter(resource => resource.resource_type === 'link').length;

            sourceProjectSummary.innerHTML = `
                <div class="font-semibold text-slate-300">${escapeHtml(source.name)}</div>
                <div class="mt-1 text-[10px] text-slate-600">
                    ${resources.length} ferramenta${resources.length === 1 ? '' : 's'} •
                    ${apps} Apps + Usados • ${locals} Apps Locais • ${links} Links
                </div>
            `;

            runImportProjectToolsBtn.disabled = resources.length === 0;
        }

        function openImportProjectToolsModal() {
            const current = selectedProject();
            if (!current) return;

            const sources = state.projects.filter(
                project => Number(project.id) !== Number(current.id)
            );

            sourceProjectSelect.innerHTML = sources.length
                ? sources.map(project => `
                    <option value="${Number(project.id)}">${escapeHtml(project.name)}</option>
                `).join('')
                : '<option value="">Nenhum outro projeto disponível</option>';

            importProjectToolsError.classList.add('hidden');
            importProjectToolsError.textContent = '';
            runImportProjectToolsBtn.disabled = !sources.length;

            importProjectToolsModal.classList.remove('hidden');
            importProjectToolsModal.classList.add('flex');
            updateSourceProjectSummary();
        }

        function closeImportProjectToolsModal() {
            importProjectToolsModal.classList.add('hidden');
            importProjectToolsModal.classList.remove('flex');
            importProjectToolsError.classList.add('hidden');
        }

        async function importToolsFromProject() {
            const target = selectedProject();
            const sourceId = Number(sourceProjectSelect.value);

            if (!target || !sourceId) return;

            runImportProjectToolsBtn.disabled = true;
            runImportProjectToolsBtn.textContent = 'Importando...';
            importProjectToolsError.classList.add('hidden');

            try {
                const response = await fetch(
                    `/api/projects/${target.id}/resources/import/${sourceId}`,
                    { method: 'POST' }
                );

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível trazer as ferramentas.');
                }

                const index = state.projects.findIndex(
                    project => Number(project.id) === Number(data.project.id)
                );
                if (index !== -1) state.projects[index] = data.project;

                closeImportProjectToolsModal();
                renderProjects();
                renderSelectedProject();

                const details = [
                    `${data.imported} adicionada${data.imported === 1 ? '' : 's'}`,
                    `${data.duplicates} já existente${data.duplicates === 1 ? '' : 's'}`
                ];
                if (data.missing) {
                    details.push(`${data.missing} indisponível${data.missing === 1 ? '' : 'is'}`);
                }

                showToast(`Ferramentas importadas: ${details.join(' • ')}.`);
            } catch (error) {
                importProjectToolsError.textContent = error.message || 'Erro ao importar ferramentas.';
                importProjectToolsError.classList.remove('hidden');
            } finally {
                runImportProjectToolsBtn.disabled = false;
                runImportProjectToolsBtn.textContent = 'Trazer ferramentas';
            }
        }

        async function addResourceToProject(resourceType, resourceId) {
            const project = selectedProject();
            if (!project) return;

            try {
                const response = await fetch(`/api/projects/${project.id}/resources`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        resource_type: resourceType,
                        resource_id: resourceType === 'local_app'
                            ? String(resourceId)
                            : Number(resourceId)
                    })
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível adicionar o recurso.');
                }

                const index = state.projects.findIndex(item => Number(item.id) === Number(data.id));
                if (index !== -1) state.projects[index] = data;

                renderProjects();
                renderSelectedProject();
                showToast('Ferramenta adicionada ao projeto.');
            } catch (error) {
                showToast(error.message || 'Erro ao adicionar ferramenta.', true);
            }
        }

        async function removeResourceFromProject(resourceType, resourceId) {
            const project = selectedProject();
            if (!project) return;

            try {
                const response = await fetch(
                    `/api/projects/${project.id}/resources/${encodeURIComponent(resourceType)}/${encodeURIComponent(String(resourceId))}`,
                    { method: 'DELETE' }
                );

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Não foi possível remover o recurso.');
                }

                const index = state.projects.findIndex(item => Number(item.id) === Number(data.id));
                if (index !== -1) state.projects[index] = data;

                renderProjects();
                renderSelectedProject();
            } catch (error) {
                showToast(error.message || 'Erro ao remover ferramenta.', true);
            }
        }

        workspaceExplorerList.addEventListener('click', event => {
            const actionButton = event.target.closest('[data-workspace-action]');
            if (actionButton) {
                event.stopPropagation();
                const action = actionButton.dataset.workspaceAction;
                const folderName = actionButton.dataset.folder;

                if (action === 'details') {
                    state.selectedFolderName = folderName;
                    state.selectedProjectId = Number(actionButton.dataset.projectId);
                    workspaceExplorerSelection.textContent = `Selecionado: ${folderName}`;
                    renderWorkspaceExplorer();
                    renderProjects();
                    renderSelectedProject();
                    projectArea.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    return;
                }

                if (action === 'register') {
                    state.selectedFolderName = folderName;
                    openProjectModal(null, folderName);
                    return;
                }

                runWorkspaceFolderAction(folderName, action);
                return;
            }

            const row = event.target.closest('[data-workspace-folder]');
            if (!row) return;

            state.selectedFolderName = row.dataset.workspaceFolder || '';
            state.selectedProjectId = Number(row.dataset.projectId || 0) || null;
            workspaceExplorerSelection.textContent = state.selectedFolderName
                ? `Selecionado: ${state.selectedFolderName}`
                : 'Nenhum projeto selecionado';

            renderWorkspaceExplorer();
            renderProjects();
            renderSelectedProject();
        });

        workspaceExplorerList.addEventListener('dblclick', event => {
            if (event.target.closest('button, a')) return;
            const row = event.target.closest('[data-workspace-folder]');
            if (row) runWorkspaceFolderAction(row.dataset.workspaceFolder, 'open');
        });

        document.getElementById('workspaceExplorerRefreshBtn').addEventListener('click', async () => {
            try {
                await loadWorkspaceExplorer();
                showToast('Lista de pastas atualizada.');
            } catch (error) {
                showToast(error.message || 'Erro ao atualizar Workspace.', true);
            }
        });

        document.getElementById('workspaceExplorerNewBtn').addEventListener('click', () => openProjectModal());

        workspaceExplorerSearch.addEventListener('input', () => {
            state.workspaceSearch = workspaceExplorerSearch.value;
            renderWorkspaceExplorer();
        });

        document.querySelectorAll('[data-workspace-filter]').forEach(button => {
            button.addEventListener('click', () => {
                state.workspaceFilter = button.dataset.workspaceFilter;
                renderWorkspaceExplorer();
            });
        });

        projectsList.addEventListener('click', event => {
            const button = event.target.closest('[data-project-id]');
            if (!button) return;
            state.selectedProjectId = Number(button.dataset.projectId);
            renderProjects();
            renderSelectedProject();
        });

        document.getElementById('newProjectBtn').addEventListener('click', () => openProjectModal());
        document.getElementById('newProjectMiniBtn').addEventListener('click', () => openProjectModal());
        document.getElementById('editProjectBtn').addEventListener('click', () => {
            const project = selectedProject();
            if (project) openProjectModal(project);
        });
        document.getElementById('deleteProjectBtn').addEventListener('click', deleteSelectedProject);

        document.getElementById('newPhaseBtn').addEventListener('click', () => openPhaseModal());
        document.getElementById('phaseEmptyAddBtn').addEventListener('click', () => openPhaseModal());

        phaseTimeline.addEventListener('click', event => {
            const editButton = event.target.closest('[data-edit-phase]');
            if (editButton) {
                const project = selectedProject();
                const phase = (project?.phases || []).find(
                    item => Number(item.id) === Number(editButton.dataset.editPhase)
                );
                if (phase) openPhaseModal(phase);
                return;
            }

            const deleteButton = event.target.closest('[data-delete-phase]');
            if (deleteButton) {
                deletePhase(Number(deleteButton.dataset.deletePhase));
            }
        });

        phaseProgressInput.addEventListener('input', () => {
            phaseProgressValue.textContent = `${phaseProgressInput.value}%`;
        });

        phaseStatusInput.addEventListener('change', () => {
            if (phaseStatusInput.value === 'concluido') {
                phaseProgressInput.value = 100;
                phaseProgressValue.textContent = '100%';
            }
        });

        document.getElementById('closePhaseModalBtn').addEventListener('click', closePhaseModal);
        document.getElementById('cancelPhaseBtn').addEventListener('click', closePhaseModal);
        phaseForm.addEventListener('submit', savePhase);
        phaseModal.addEventListener('click', event => {
            if (event.target === phaseModal) closePhaseModal();
        });

        document.getElementById('closeProjectModalBtn').addEventListener('click', closeProjectModal);
        document.getElementById('cancelProjectBtn').addEventListener('click', closeProjectModal);
        projectForm.addEventListener('submit', saveProject);
        projectModal.addEventListener('click', event => {
            if (event.target === projectModal) closeProjectModal();
        });

        document.querySelectorAll('[data-library-tab]').forEach(button => {
            button.addEventListener('click', () => {
                state.libraryTab = button.dataset.libraryTab;

                document.querySelectorAll('[data-library-tab]').forEach(tab => {
                    const active = tab.dataset.libraryTab === state.libraryTab;
                    tab.className = active
                        ? 'library-tab min-h-9 rounded-lg border border-cyan-400/25 bg-cyan-400/10 px-2 py-2 text-[10px] font-bold text-cyan-200 sm:text-xs'
                        : 'library-tab min-h-9 rounded-lg border border-slate-700/25 bg-[#071722] px-2 py-2 text-[10px] font-bold text-slate-500 sm:text-xs';
                });

                resourceSearch.value = '';
                populateResourceCategoryFilter(true);
                renderLibrary();
            });
        });

        resourceSearch.addEventListener('input', renderLibrary);
        resourceCategoryFilter.addEventListener('change', renderLibrary);

        resourceLibrary.addEventListener('dragstart', event => {
            const card = event.target.closest('[data-library-resource-type]');
            if (!card) return;

            card.classList.add('dragging');

            if (event.dataTransfer) {
                event.dataTransfer.effectAllowed = 'copy';
                event.dataTransfer.setData(
                    'application/x-tech-tool-resource',
                    JSON.stringify({
                        resource_type: card.dataset.libraryResourceType,
                        resource_id: card.dataset.libraryResourceType === 'local_app'
                            ? String(card.dataset.libraryResourceId)
                            : Number(card.dataset.libraryResourceId)
                    })
                );
            }
        });

        resourceLibrary.addEventListener('dragend', event => {
            event.target.closest('[data-library-resource-type]')?.classList.remove('dragging');
        });

        resourceLibrary.addEventListener('click', event => {
            const addButton = event.target.closest('[data-add-resource-type]');
            if (!addButton) return;

            addResourceToProject(
                addButton.dataset.addResourceType,
                addButton.dataset.addResourceType === 'local_app'
                    ? String(addButton.dataset.addResourceId)
                    : Number(addButton.dataset.addResourceId)
            );
        });

        projectDropZone.addEventListener('dragover', event => {
            if (!selectedProject()) return;
            event.preventDefault();
            if (event.dataTransfer) event.dataTransfer.dropEffect = 'copy';
            projectDropZone.classList.add('drag-over');
        });

        projectDropZone.addEventListener('dragleave', event => {
            if (!projectDropZone.contains(event.relatedTarget)) {
                projectDropZone.classList.remove('drag-over');
            }
        });

        projectDropZone.addEventListener('drop', event => {
            event.preventDefault();
            projectDropZone.classList.remove('drag-over');

            const raw = event.dataTransfer?.getData('application/x-tech-tool-resource');
            if (!raw) return;

            try {
                const resource = JSON.parse(raw);
                addResourceToProject(resource.resource_type, resource.resource_id);
            } catch {
                showToast('Não foi possível identificar a ferramenta arrastada.', true);
            }
        });

        projectResources.addEventListener('click', event => {
            const launchButton = event.target.closest('[data-launch-local-id]');
            if (launchButton && !launchButton.disabled) {
                launchWorkspaceLocalApp(launchButton.dataset.launchLocalId);
                return;
            }

            const button = event.target.closest('[data-remove-resource-type]');
            if (!button) return;

            removeResourceFromProject(
                button.dataset.removeResourceType,
                button.dataset.removeResourceType === 'local_app'
                    ? String(button.dataset.removeResourceId)
                    : Number(button.dataset.removeResourceId)
            );
        });

        importProjectToolsBtn.addEventListener('click', openImportProjectToolsModal);
        document.getElementById('closeImportProjectToolsBtn').addEventListener('click', closeImportProjectToolsModal);
        document.getElementById('cancelImportProjectToolsBtn').addEventListener('click', closeImportProjectToolsModal);
        runImportProjectToolsBtn.addEventListener('click', importToolsFromProject);
        sourceProjectSelect.addEventListener('change', updateSourceProjectSummary);
        importProjectToolsModal.addEventListener('click', event => {
            if (event.target === importProjectToolsModal) closeImportProjectToolsModal();
        });

        const workspaceSidebar = document.getElementById('workspaceSidebar');
        const workspaceSidebarOverlay = document.getElementById('workspaceSidebarOverlay');

        document.getElementById('workspaceMenuBtn').addEventListener('click', () => {
            workspaceSidebar.classList.add('open');
            workspaceSidebarOverlay.classList.remove('hidden');
        });

        workspaceSidebarOverlay.addEventListener('click', () => {
            workspaceSidebar.classList.remove('open');
            workspaceSidebarOverlay.classList.add('hidden');
        });

        document.addEventListener('keydown', event => {
            if (event.key === 'Escape') {
                closeProjectModal();
                closePhaseModal();
                closeImportProjectToolsModal();
            }
        });

        loadWorkspace();
    </script>
</body>
</html>
"""



# ============================================================
# FRONTEND - GIT & GITHUB
# ============================================================

GIT_GITHUB_HTML = r"""
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="dark">
    <link rel="icon" type="image/png" href="/brand-image">
    <title>Git & GitHub • TECH TOOL HUB</title>
    <script src="https://cdn.tailwindcss.com"></script>

    <style>
        body {
            min-height: 100vh;
            margin: 0;
            background:
                radial-gradient(circle at 78% 4%, rgba(44,196,219,.08), transparent 28rem),
                linear-gradient(145deg, #05101a 0%, #071722 48%, #05111c 100%);
        }

        * { scrollbar-width: thin; scrollbar-color: #1a4055 #071722; }

        .glass-top {
            background: rgba(7,24,36,.92);
            backdrop-filter: blur(18px);
            -webkit-backdrop-filter: blur(18px);
        }

        .sidebar-item {
            transition: border-color .18s ease, background .18s ease, color .18s ease;
        }

        .sidebar-item:hover,
        .sidebar-item.active {
            color: #dffaff;
            background: linear-gradient(90deg, rgba(28,141,169,.27), rgba(14,72,94,.13));
            border-color: rgba(53,213,230,.20);
        }

        .sidebar-item.active::before {
            content: '';
            position: absolute;
            left: 0;
            top: 9px;
            bottom: 9px;
            width: 3px;
            border-radius: 0 4px 4px 0;
            background: #35d5e6;
        }

        .panel {
            border: 1px solid rgba(129,180,204,.13);
            background: linear-gradient(145deg, rgba(13,35,50,.93), rgba(8,25,37,.96));
        }

        .status-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(5,22,33,.62);
        }

        .action-card {
            border: 1px solid rgba(129,180,204,.12);
            background: rgba(6,25,37,.76);
        }

        .log-box {
            background: #04111a;
            border: 1px solid rgba(129,180,204,.12);
        }

        .finding-high {
            border-color: rgba(248,113,113,.25);
            background: rgba(127,29,29,.12);
        }

        .finding-warning {
            border-color: rgba(251,191,36,.22);
            background: rgba(120,53,15,.10);
        }


        .repo-card {
            border: 1px solid rgba(129,180,204,.12);
            background: linear-gradient(145deg, rgba(7,28,41,.90), rgba(5,20,31,.94));
            transition: transform .16s ease, border-color .16s ease, background .16s ease;
        }

        .repo-card:hover {
            transform: translateY(-1px);
            border-color: rgba(53,213,230,.22);
            background: linear-gradient(145deg, rgba(9,34,48,.95), rgba(6,23,34,.97));
        }

        .modal-shell {
            background: rgba(1, 9, 15, .76);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
        }

        @media (max-width: 1023px) {
            #gitSidebar {
                transform: translateX(-100%);
                transition: transform .22s ease;
            }

            #gitSidebar.open {
                transform: translateX(0);
            }
        }
    </style>
</head>

<body class="text-slate-100 antialiased">
    <div id="gitSidebarOverlay" class="fixed inset-0 z-30 hidden bg-black/55 lg:hidden"></div>

    <aside id="gitSidebar" class="fixed inset-y-0 left-0 z-40 flex w-[242px] flex-col border-r border-slate-700/20 bg-[#071722]/95 shadow-2xl lg:translate-x-0">
        <a href="/" class="flex h-[132px] flex-col items-center justify-center border-b border-slate-700/20 px-3 py-3 text-center">
            <img src="/brand-image" alt="TECH TOOL HUB" class="max-h-[82px] w-full max-w-[104px] object-contain">
            <div class="mt-2 leading-tight">
                <div class="text-sm font-black tracking-[.14em] text-slate-100">TECH TOOL HUB</div>
                <div class="mt-1 text-[10px] font-semibold uppercase tracking-[.30em] text-cyan-400/75">LOCAL WORKSPACE</div>
            </div>
        </a>

        <nav class="flex-1 overflow-y-auto px-3 py-5">
            <p class="mb-2 px-3 text-[10px] font-bold uppercase tracking-[.18em] text-slate-600">Navegação</p>

            <a href="/apps" class="sidebar-item relative flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▦</span><span class="font-semibold">APPS + USADOS</span>
            </a>
            <a href="/local-apps" class="sidebar-item relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">▣</span><span class="font-semibold">APPS LOCAIS</span>
            </a>
            <a href="/links" class="sidebar-item relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">★</span><span class="font-semibold">LINKS FAVORITOS</span>
            </a>
            <a href="/workspace" class="sidebar-item relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">◈</span><span class="font-semibold">PERSONALIZAR ÁREA</span>
            </a>
            <a href="/git-github" class="sidebar-item active relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-300">
                <span class="w-5 text-center">⑂</span><span class="font-semibold">GIT &amp; GITHUB</span>
            </a>
            <a href="/windows-diagnostics" class="sidebar-item relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">🩺</span><span class="font-semibold">DIAGNÓSTICO WIN</span>
            </a>
            <a href="/" class="sidebar-item relative mt-1 flex items-center gap-3 rounded-lg border border-transparent px-3 py-2.5 text-sm text-slate-400">
                <span class="w-5 text-center">☷</span><span class="font-semibold">All tools</span>
            </a>
        </nav>
    </aside>

    <div class="min-h-screen lg:pl-[242px]">
        <header class="glass-top sticky top-0 z-20 border-b border-slate-700/20">
            <div class="flex min-h-[72px] flex-wrap items-center gap-3 px-4 py-3 sm:px-6 xl:px-8">
                <button id="gitMenuBtn" class="grid h-10 w-10 place-items-center rounded-xl border border-slate-700/30 bg-[#0a1d2b] text-slate-300 lg:hidden">☰</button>

                <div class="min-w-[200px] flex-1">
                    <h1 class="text-base font-bold tracking-[.12em] text-slate-100">GIT &amp; GITHUB</h1>
                    <p class="mt-0.5 text-[10px] uppercase tracking-[.17em] text-slate-600">Segurança • Commit • Push • Release</p>
                </div>

                <div class="flex flex-wrap items-center gap-2">
                    <button id="gitCreateExecutableBtn" type="button" class="h-10 rounded-xl border border-cyan-400/20 bg-cyan-400/10 px-4 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">
                        ⚙ Criar Executável
                    </button>
                    <button id="gitShowReadmeBtn" type="button" class="h-10 rounded-xl border border-slate-700/30 bg-[#091c29] px-4 text-xs font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">
                        ▤ Readme
                    </button>
                    <button id="refreshGitStatusBtn" class="h-10 rounded-xl border border-slate-700/30 bg-[#091c29] px-4 text-xs font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">
                        ↻ Atualizar status
                    </button>
                </div>
            </div>
        </header>

        <main class="mx-auto max-w-[1500px] px-4 py-6 sm:px-6 xl:px-8">
            <section class="mb-5">
                <p class="text-xs font-semibold uppercase tracking-[.18em] text-cyan-400/65">Pipeline local seguro</p>
                <h2 class="mt-1 text-2xl font-bold tracking-tight text-slate-100">Publicação do projeto</h2>
                <p class="mt-1 max-w-4xl text-sm leading-6 text-slate-500">
                    Audite vazamentos, publique alterações no GitHub e crie releases sem gravar token no TECH TOOL HUB.
                    O acesso ao GitHub usa a autenticação existente do Git/GitHub CLI.
                </p>

                <div class="mt-3 rounded-xl border border-emerald-400/20 bg-emerald-400/[.035] px-4 py-3 text-xs leading-5 text-slate-500">
                    <strong class="text-emerald-200">🔒 Zero credenciais:</strong>
                    o Hub não armazena senhas, tokens, cookies ou sessões. O GitHub usa exclusivamente a sessão já autenticada do <code class="text-slate-300">gh</code>.
                </div>
            </section>


            <!-- GERENCIADOR DE REPOSITÓRIOS GITHUB -->
            <section class="panel mb-5 rounded-2xl p-4 sm:p-5">
                <div class="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
                    <div>
                        <p class="text-[10px] font-bold uppercase tracking-[.14em] text-cyan-400/60">Conta GitHub</p>
                        <h3 class="mt-1 text-lg font-bold text-slate-100">Meus repositórios</h3>
                        <p class="mt-1 max-w-3xl text-xs leading-5 text-slate-500">
                            Consulte, filtre, clone, edite e administre seus repositórios sem sair do TECH TOOL HUB.
                            As operações usam sua sessão autenticada do GitHub CLI.
                        </p>
                    </div>
                    <div class="flex flex-wrap gap-2">
                        <button id="createGithubRepoBtn" class="h-10 rounded-xl border border-emerald-400/20 bg-emerald-400/10 px-4 text-xs font-bold text-emerald-200 hover:bg-emerald-400/15">+ Novo repositório</button>
                        <button id="refreshGithubReposBtn" class="h-10 rounded-xl border border-cyan-400/20 bg-cyan-400/10 px-4 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">↻ Atualizar lista</button>
                    </div>
                </div>

                <div id="githubRepoAccountMessage" class="mt-3 text-[10px] text-slate-600">Carregando conta...</div>

                <div class="mt-4 grid grid-cols-2 gap-2 sm:grid-cols-3 xl:grid-cols-6">
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountTotal" class="text-xl font-black text-slate-100">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">Total</div></div>
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountPublic" class="text-xl font-black text-cyan-200">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">Públicos</div></div>
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountPrivate" class="text-xl font-black text-violet-200">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">Privados</div></div>
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountLocal" class="text-xl font-black text-emerald-200">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">No PC</div></div>
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountArchived" class="text-xl font-black text-amber-200">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">Arquivados</div></div>
                    <div class="status-card rounded-xl p-3 text-center"><div id="repoCountForks" class="text-xl font-black text-slate-300">—</div><div class="text-[9px] uppercase tracking-[.08em] text-slate-600">Forks</div></div>
                </div>

                <div class="mt-4 grid gap-2 lg:grid-cols-[minmax(0,1fr)_190px_190px_170px]">
                    <input id="githubRepoSearchInput" placeholder="Pesquisar por nome, descrição ou linguagem..." class="h-11 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                    <select id="githubRepoVisibilityFilter" class="h-11 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-300 outline-none">
                        <option value="all">Todas visibilidades</option><option value="public">Públicos</option><option value="private">Privados</option><option value="internal">Internos</option>
                    </select>
                    <select id="githubRepoStateFilter" class="h-11 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-300 outline-none">
                        <option value="all">Todos estados</option><option value="active">Ativos</option><option value="archived">Arquivados</option><option value="local">Clonados neste PC</option><option value="fork">Forks</option>
                    </select>
                    <select id="githubRepoSortSelect" class="h-11 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-300 outline-none">
                        <option value="updated">Mais recentes</option><option value="name">Nome A-Z</option><option value="stars">Mais estrelas</option><option value="size">Maior tamanho</option>
                    </select>
                </div>

                <div class="mt-3 flex flex-wrap items-center gap-3 text-[10px] text-slate-600">
                    <label class="flex items-center gap-2">
                        <span>Escopo</span>
                        <select id="githubRepoScopeSelect" class="h-8 rounded-lg border border-slate-700/30 bg-[#061722] px-2 text-[10px] text-slate-300 outline-none">
                            <option value="owner">Meus repositórios</option>
                            <option value="all">Todos a que tenho acesso</option>
                        </select>
                    </label>
                    <span id="githubRepoFilteredCount">—</span>
                </div>

                <div id="githubReposLoading" class="mt-5 hidden rounded-xl border border-slate-700/20 bg-[#061722] p-5 text-center text-xs text-slate-500">Consultando GitHub...</div>
                <div id="githubReposError" class="mt-5 hidden rounded-xl border border-red-900/40 bg-red-950/20 p-4 text-xs leading-5 text-red-300"></div>
                <div id="githubRepoGrid" class="mt-5 grid gap-3 md:grid-cols-2 2xl:grid-cols-3"></div>
            </section>

            <!-- SELETOR DE PROJETO -->
            <section class="panel mb-5 rounded-2xl p-4 sm:p-5">
                <div class="flex flex-col gap-4 xl:flex-row xl:items-end">
                    <div class="min-w-0 flex-1">
                        <div class="flex flex-wrap items-center gap-2">
                            <h3 class="text-sm font-bold text-slate-200">Projeto ativo</h3>
                            <span id="selectedProjectGitBadge" class="rounded-md border border-slate-700/25 px-2 py-0.5 text-[9px] font-bold uppercase tracking-[.08em] text-slate-500">...</span>
                        </div>
                        <p class="mt-1 text-[10px] leading-5 text-slate-600">
                            Selecione qualquer repositório do computador. A checagem, commit, push e Release serão executados somente no projeto ativo.
                        </p>

                        <div class="mt-3 grid gap-2 lg:grid-cols-[minmax(0,1fr)_auto_auto]">
                            <select id="projectSelect" class="h-11 min-w-0 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                                <option>Carregando projetos...</option>
                            </select>

                            <button id="chooseProjectFolderBtn" class="h-11 rounded-xl border border-cyan-400/20 bg-cyan-400/10 px-4 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15">
                                📁 Selecionar pasta
                            </button>

                            <button id="forgetProjectBtn" class="h-11 rounded-xl border border-slate-700/25 bg-[#091c29] px-4 text-xs font-bold text-slate-400 hover:text-slate-200">
                                Remover da lista
                            </button>
                        </div>

                        <div id="selectedProjectPath" class="mt-2 break-all text-[10px] text-slate-600">—</div>
                    </div>
                </div>
            </section>

            <!-- STATUS -->
            <section class="panel rounded-2xl p-4 sm:p-5">
                <div class="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                        <h3 class="text-sm font-bold text-slate-200">Ambiente</h3>
                        <p id="environmentMessage" class="mt-0.5 text-[10px] text-slate-600">Verificando Git e GitHub...</p>
                    </div>
                    <span id="sourceModeBadge" class="rounded-lg border border-slate-700/25 px-2.5 py-1 text-[10px] font-bold uppercase tracking-[.08em] text-slate-500">...</span>
                </div>

                <div class="mt-4 grid grid-cols-2 gap-3 lg:grid-cols-4">
                    <div class="status-card rounded-xl p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Branch</div>
                        <div id="gitBranch" class="mt-1 truncate text-sm font-bold text-slate-300">—</div>
                    </div>
                    <div class="status-card rounded-xl p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Alterações locais</div>
                        <div id="gitChangedFiles" class="mt-1 text-xl font-black text-slate-200">0</div>
                    </div>
                    <div class="status-card rounded-xl p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">GitHub CLI</div>
                        <div id="ghStatus" class="mt-1 text-sm font-bold text-slate-300">—</div>
                    </div>
                    <div class="status-card rounded-xl p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Artefato de Release</div>
                        <div id="latestReleaseName" class="mt-1 truncate text-xs font-bold text-slate-300">—</div>
                    </div>
                </div>

                <div class="mt-3 rounded-xl border border-slate-700/20 bg-[#061722] px-4 py-3">
                    <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Origin</div>
                    <div id="gitRemote" class="mt-1 break-all text-xs text-slate-400">—</div>
                </div>
            </section>

            <!-- AÇÕES -->
            <div class="mt-5 grid gap-5 xl:grid-cols-3">
                <!-- SEGURANÇA -->
                <section class="action-card rounded-2xl p-5">
                    <div class="flex items-start justify-between gap-3">
                        <div>
                            <div class="text-[10px] font-bold uppercase tracking-[.12em] text-cyan-400/60">Etapa 1</div>
                            <h3 class="mt-1 text-lg font-bold text-slate-100">Checar vazamentos</h3>
                            <p class="mt-1 text-xs leading-5 text-slate-500">Verifica arquivos que podem entrar no Git e procura credenciais, dados locais e artefatos proibidos.</p>
                        </div>
                        <div class="text-2xl">⌕</div>
                    </div>

                    <div class="mt-4 grid grid-cols-3 gap-2">
                        <div class="rounded-lg border border-slate-700/20 bg-[#061722] p-2 text-center">
                            <div id="auditCandidates" class="text-lg font-black text-slate-200">—</div>
                            <div class="text-[9px] text-slate-600">arquivos</div>
                        </div>
                        <div class="rounded-lg border border-slate-700/20 bg-[#061722] p-2 text-center">
                            <div id="auditBlocking" class="text-lg font-black text-red-300">—</div>
                            <div class="text-[9px] text-slate-600">riscos</div>
                        </div>
                        <div class="rounded-lg border border-slate-700/20 bg-[#061722] p-2 text-center">
                            <div id="auditWarnings" class="text-lg font-black text-amber-200">—</div>
                            <div class="text-[9px] text-slate-600">avisos</div>
                        </div>
                    </div>

                    <button id="runSecurityAuditBtn" class="mt-4 h-11 w-full rounded-xl border border-cyan-400/20 bg-cyan-400/10 text-xs font-bold text-cyan-200 hover:bg-cyan-400/15 disabled:cursor-not-allowed disabled:opacity-40">
                        Executar checagem
                    </button>

                    <div id="auditResult" class="mt-4 hidden"></div>
                </section>

                <!-- GIT PUSH -->
                <section class="action-card rounded-2xl p-5">
                    <div class="flex items-start justify-between gap-3">
                        <div>
                            <div class="text-[10px] font-bold uppercase tracking-[.12em] text-cyan-400/60">Etapa 2</div>
                            <h3 class="mt-1 text-lg font-bold text-slate-100">Enviar ao GitHub</h3>
                            <p class="mt-1 text-xs leading-5 text-slate-500">Executa auditoria, commit, fetch, rebase seguro quando necessário e push para origin.</p>
                        </div>
                        <div class="text-2xl">⑂</div>
                    </div>

                    <label class="mt-4 block">
                        <span class="mb-1.5 block text-[10px] font-bold uppercase tracking-[.08em] text-slate-600">Mensagem do commit</span>
                        <input id="commitMessageInput" maxlength="160" value="Atualiza projeto" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                    </label>

                    <button id="publishGitBtn" class="mt-4 h-11 w-full rounded-xl border border-emerald-400/20 bg-emerald-400/10 text-xs font-bold text-emerald-200 hover:bg-emerald-400/15 disabled:cursor-not-allowed disabled:opacity-40">
                        Auditar e publicar
                    </button>

                    <div id="publishResult" class="mt-4 hidden"></div>
                </section>

                <!-- RELEASE -->
                <section class="action-card rounded-2xl p-5">
                    <div class="flex items-start justify-between gap-3">
                        <div>
                            <div class="text-[10px] font-bold uppercase tracking-[.12em] text-cyan-400/60">Etapa 3</div>
                            <h3 class="mt-1 text-lg font-bold text-slate-100">Criar Release</h3>
                            <p class="mt-1 text-xs leading-5 text-slate-500">Cria uma Release do projeto ativo e pode anexar um artefato da pasta release.</p>
                        </div>
                        <div class="text-2xl">⬡</div>
                    </div>

                    <div class="mt-4 grid gap-3 sm:grid-cols-2 xl:grid-cols-1 2xl:grid-cols-2">
                        <label>
                            <span class="mb-1.5 block text-[10px] font-bold uppercase tracking-[.08em] text-slate-600">Tag</span>
                            <input id="releaseTagInput" placeholder="v2.24.0" class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                        </label>
                        <label>
                            <span class="mb-1.5 block text-[10px] font-bold uppercase tracking-[.08em] text-slate-600">Título</span>
                            <input id="releaseTitleInput" placeholder="TECH TOOL HUB v2.24.0" class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                        </label>
                    </div>

                    <label class="mt-3 block">
                        <span class="mb-1.5 block text-[10px] font-bold uppercase tracking-[.08em] text-slate-600">Artefato da pasta release</span>
                        <select id="releaseAssetSelect" class="h-10 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35">
                            <option value="">Sem artefato / detectar automaticamente</option>
                        </select>
                        <p class="mt-1 text-[9px] leading-4 text-slate-600">
                            O painel prioriza o executável principal e ignora arquivos internos do PyInstaller, como base_library.zip e _internal.
                        </p>
                    </label>

                    <label class="mt-3 block">
                        <span class="mb-1.5 block text-[10px] font-bold uppercase tracking-[.08em] text-slate-600">Notas</span>
                        <textarea id="releaseNotesInput" rows="4" maxlength="5000" placeholder="Principais mudanças desta versão..." class="w-full resize-none rounded-xl border border-slate-700/30 bg-[#061722] px-3 py-3 text-sm text-slate-200 outline-none focus:border-cyan-400/35"></textarea>
                    </label>

                    <div class="mt-3 flex flex-wrap gap-4 text-xs text-slate-500">
                        <label class="flex items-center gap-2">
                            <input id="includeCertificateInput" type="checkbox" checked class="accent-cyan-400">
                            Incluir .cer
                        </label>
                        <label class="flex items-center gap-2">
                            <input id="releaseDraftInput" type="checkbox" class="accent-cyan-400">
                            Criar como rascunho
                        </label>
                    </div>

                    <button id="createReleaseBtn" class="mt-4 h-11 w-full rounded-xl border border-blue-400/20 bg-blue-400/10 text-xs font-bold text-blue-200 hover:bg-blue-400/15 disabled:cursor-not-allowed disabled:opacity-40">
                        Criar GitHub Release
                    </button>

                    <div id="releaseResult" class="mt-4 hidden"></div>
                </section>
            </div>


            <!-- COMANDOS INDIVIDUAIS GIT -->
            <section class="panel mt-5 rounded-2xl p-5">
                <div class="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
                    <div>
                        <div class="text-[10px] font-bold uppercase tracking-[.14em] text-cyan-400/60">Ferramentas avançadas</div>
                        <h3 class="mt-1 text-lg font-bold text-slate-100">Comandos individuais Git</h3>
                        <p class="mt-1 max-w-4xl text-xs leading-5 text-slate-500">
                            Execute operações isoladas no projeto ativo. O Dashboard usa uma lista fechada de comandos:
                            não executa texto arbitrário no PowerShell e mantém as proteções de vazamento em Add, Commit e Push.
                        </p>
                    </div>
                    <span id="gitIndividualRepoBadge" class="rounded-lg border border-slate-700/25 bg-[#061722] px-3 py-1.5 text-[10px] font-bold uppercase tracking-[.08em] text-slate-500">
                        aguardando projeto
                    </span>
                </div>

                <!-- LEITURA / SINCRONIZAÇÃO -->
                <div class="mt-5">
                    <div class="mb-2 text-[10px] font-bold uppercase tracking-[.12em] text-slate-600">Consulta e sincronização</div>
                    <div class="grid grid-cols-2 gap-2 sm:grid-cols-4 xl:grid-cols-8">
                        <button data-simple-git-action="status" class="git-repo-command h-10 rounded-xl border border-slate-700/25 bg-[#071c29] px-3 text-[11px] font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">Status</button>
                        <button data-simple-git-action="branches" class="git-repo-command h-10 rounded-xl border border-slate-700/25 bg-[#071c29] px-3 text-[11px] font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">Branches</button>
                        <button data-simple-git-action="log" class="git-repo-command h-10 rounded-xl border border-slate-700/25 bg-[#071c29] px-3 text-[11px] font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">Log</button>
                        <button data-simple-git-action="diff" class="git-repo-command h-10 rounded-xl border border-slate-700/25 bg-[#071c29] px-3 text-[11px] font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">Diff</button>
                        <button data-simple-git-action="remote" class="git-repo-command h-10 rounded-xl border border-slate-700/25 bg-[#071c29] px-3 text-[11px] font-bold text-slate-300 hover:border-cyan-400/25 hover:text-cyan-200">Remote</button>
                        <button data-simple-git-action="fetch" class="git-repo-command h-10 rounded-xl border border-blue-400/20 bg-blue-400/5 px-3 text-[11px] font-bold text-blue-200 hover:bg-blue-400/10">Fetch</button>
                        <button data-simple-git-action="pull_rebase" class="git-repo-command h-10 rounded-xl border border-blue-400/20 bg-blue-400/5 px-3 text-[11px] font-bold text-blue-200 hover:bg-blue-400/10">Pull --rebase</button>
                        <button data-simple-git-action="push" class="git-repo-command h-10 rounded-xl border border-emerald-400/20 bg-emerald-400/5 px-3 text-[11px] font-bold text-emerald-200 hover:bg-emerald-400/10">Push</button>
                    </div>
                </div>

                <div class="mt-5 grid gap-4 xl:grid-cols-2 2xl:grid-cols-4">
                    <!-- INIT / CLONE -->
                    <div class="action-card rounded-xl p-4">
                        <div class="flex items-center justify-between">
                            <h4 class="text-sm font-bold text-slate-200">Repositório</h4>
                            <span class="text-lg">⌘</span>
                        </div>
                        <p class="mt-1 text-[10px] leading-4 text-slate-600">Inicialize a pasta ativa ou clone outro repositório.</p>

                        <button id="gitInitBtn" class="mt-3 h-10 w-full rounded-lg border border-cyan-400/20 bg-cyan-400/5 text-[11px] font-bold text-cyan-200 hover:bg-cyan-400/10">
                            git init
                        </button>

                        <div class="mt-4 border-t border-slate-700/20 pt-4">
                            <label class="block">
                                <span class="mb-1 block text-[9px] font-bold uppercase tracking-[.08em] text-slate-600">URL HTTPS / SSH</span>
                                <input id="gitCloneUrlInput" placeholder="https://github.com/usuario/projeto.git" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">
                            </label>

                            <div class="mt-2 grid grid-cols-[minmax(0,1fr)_auto] gap-2">
                                <input id="gitCloneParentInput" placeholder="Pasta-pai de destino" class="h-10 min-w-0 rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">
                                <button id="browseCloneParentBtn" class="h-10 rounded-lg border border-slate-700/30 bg-[#091c29] px-3 text-[11px] font-bold text-slate-300 hover:text-cyan-200">📁</button>
                            </div>

                            <input id="gitCloneFolderInput" placeholder="Nome da pasta (opcional)" class="mt-2 h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">

                            <button id="gitCloneBtn" class="mt-2 h-10 w-full rounded-lg border border-blue-400/20 bg-blue-400/5 text-[11px] font-bold text-blue-200 hover:bg-blue-400/10">
                                git clone
                            </button>
                        </div>
                    </div>

                    <!-- ADD / COMMIT -->
                    <div class="action-card rounded-xl p-4">
                        <div class="flex items-center justify-between">
                            <h4 class="text-sm font-bold text-slate-200">Stage &amp; Commit</h4>
                            <span class="text-lg">＋</span>
                        </div>
                        <p class="mt-1 text-[10px] leading-4 text-slate-600">Add e Commit passam pela auditoria de segurança.</p>

                        <label class="mt-3 block">
                            <span class="mb-1 block text-[9px] font-bold uppercase tracking-[.08em] text-slate-600">Arquivo/pasta para Add</span>
                            <input id="gitAddPathInput" placeholder="vazio = git add -A" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">
                        </label>

                        <button id="gitAddBtn" class="mt-2 h-10 w-full rounded-lg border border-cyan-400/20 bg-cyan-400/5 text-[11px] font-bold text-cyan-200 hover:bg-cyan-400/10">
                            git add
                        </button>

                        <label class="mt-3 block">
                            <span class="mb-1 block text-[9px] font-bold uppercase tracking-[.08em] text-slate-600">Mensagem do Commit</span>
                            <input id="gitCommitMessageInput" maxlength="240" placeholder="Descrição da alteração" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">
                        </label>

                        <button id="gitCommitBtn" class="mt-2 h-10 w-full rounded-lg border border-emerald-400/20 bg-emerald-400/5 text-[11px] font-bold text-emerald-200 hover:bg-emerald-400/10">
                            git commit
                        </button>
                    </div>

                    <!-- SWITCH / BRANCH -->
                    <div class="action-card rounded-xl p-4">
                        <div class="flex items-center justify-between">
                            <h4 class="text-sm font-bold text-slate-200">Switch &amp; Branch</h4>
                            <span class="text-lg">⑂</span>
                        </div>
                        <p class="mt-1 text-[10px] leading-4 text-slate-600">Troque, crie ou remova branches locais.</p>

                        <label class="mt-3 block">
                            <span class="mb-1 block text-[9px] font-bold uppercase tracking-[.08em] text-slate-600">Branch</span>
                            <input id="gitBranchNameInput" placeholder="feature/minha-alteracao" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-cyan-400/35">
                        </label>

                        <div class="mt-2 grid grid-cols-2 gap-2">
                            <button id="gitSwitchBtn" class="h-10 rounded-lg border border-cyan-400/20 bg-cyan-400/5 text-[11px] font-bold text-cyan-200 hover:bg-cyan-400/10">Switch</button>
                            <button id="gitSwitchCreateBtn" class="h-10 rounded-lg border border-blue-400/20 bg-blue-400/5 text-[11px] font-bold text-blue-200 hover:bg-blue-400/10">Criar + Switch</button>
                        </div>

                        <button id="gitBranchCreateBtn" class="mt-2 h-10 w-full rounded-lg border border-slate-700/30 bg-[#071c29] text-[11px] font-bold text-slate-300 hover:text-cyan-200">
                            Criar branch sem trocar
                        </button>

                        <label class="mt-3 flex items-center gap-2 text-[10px] text-slate-500">
                            <input id="gitBranchForceDeleteInput" type="checkbox" class="accent-red-400">
                            Forçar exclusão (-D) se não estiver mesclada
                        </label>

                        <button id="gitBranchDeleteBtn" class="mt-2 h-10 w-full rounded-lg border border-red-900/35 bg-red-950/20 text-[11px] font-bold text-red-300 hover:bg-red-950/35">
                            Excluir branch local
                        </button>
                    </div>

                    <!-- RM CACHED -->
                    <div class="action-card rounded-xl p-4">
                        <div class="flex items-center justify-between">
                            <h4 class="text-sm font-bold text-slate-200">Remover do Git</h4>
                            <span class="text-lg">⊘</span>
                        </div>
                        <p class="mt-1 text-[10px] leading-4 text-slate-600">
                            Equivale a <strong class="font-bold text-slate-400">git rm --cached</strong>: retira do índice sem apagar do computador.
                        </p>

                        <label class="mt-3 block">
                            <span class="mb-1 block text-[9px] font-bold uppercase tracking-[.08em] text-slate-600">Arquivo ou pasta</span>
                            <input id="gitRmCachedPathInput" placeholder="release/test" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none focus:border-red-400/35">
                        </label>

                        <button id="gitRmCachedBtn" class="mt-2 h-10 w-full rounded-lg border border-red-900/35 bg-red-950/20 text-[11px] font-bold text-red-300 hover:bg-red-950/35">
                            Remover somente do índice
                        </button>

                        <div class="mt-4 rounded-lg border border-slate-700/20 bg-[#061722] p-3 text-[10px] leading-4 text-slate-600">
                            O Dashboard não disponibiliza <strong class="text-slate-400">reset --hard</strong>, limpeza forçada ou comandos shell arbitrários.
                        </div>
                    </div>
                </div>

                <div id="gitIndividualResult" class="mt-4 hidden rounded-xl border border-slate-700/25 bg-[#061722] px-4 py-3 text-xs text-slate-400"></div>
            </section>

            <!-- LOG / DETALHES -->
            <section class="panel mt-5 rounded-2xl p-5">
                <div class="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                    <div>
                        <h3 class="text-sm font-bold text-slate-200">Log da sessão</h3>
                        <p class="mt-0.5 text-[10px] text-slate-600">Saída das ações executadas nesta página.</p>
                    </div>
                    <button id="clearGitLogBtn" class="h-8 rounded-lg border border-slate-700/25 px-3 text-[10px] font-semibold text-slate-500 hover:text-slate-300">Limpar</button>
                </div>

                <pre id="gitActionLog" class="log-box mt-3 min-h-[160px] max-h-[360px] overflow-auto whitespace-pre-wrap rounded-xl p-4 text-[11px] leading-5 text-slate-500">Pronto.</pre>
            </section>
        </main>
    </div>


    <div id="githubRepoCreateModal" class="modal-shell fixed inset-0 z-[80] hidden items-center justify-center p-4">
        <div class="panel max-h-[92vh] w-full max-w-xl overflow-y-auto rounded-2xl p-5 shadow-2xl">
            <div class="flex items-start justify-between gap-4">
                <div><h3 class="text-lg font-bold text-slate-100">Novo repositório GitHub</h3><p class="mt-1 text-xs text-slate-500">Cria na conta autenticada pelo GitHub CLI.</p></div>
                <button data-close-modal="githubRepoCreateModal" class="rounded-lg px-2 py-1 text-slate-500">✕</button>
            </div>
            <div class="mt-5 grid gap-3">
                <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Nome</span><input id="newGithubRepoName" maxlength="100" placeholder="meu-projeto" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none"></label>
                <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Descrição</span><textarea id="newGithubRepoDescription" rows="3" maxlength="350" class="w-full resize-none rounded-xl border border-slate-700/30 bg-[#061722] px-3 py-3 text-sm text-slate-200 outline-none"></textarea></label>
                <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Visibilidade</span><select id="newGithubRepoVisibility" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-300 outline-none"><option value="private">Privado</option><option value="public">Público</option></select></label>
                <label class="flex items-center gap-2 text-xs text-slate-500"><input id="newGithubRepoReadme" type="checkbox" checked class="accent-cyan-400"> Inicializar com README</label>
                <label class="flex items-center gap-2 text-xs text-slate-500"><input id="newGithubRepoCloneAfter" type="checkbox" class="accent-cyan-400"> Clonar no computador após criar</label>
                <div id="newGithubRepoCloneArea" class="hidden grid grid-cols-[minmax(0,1fr)_auto] gap-2">
                    <input id="newGithubRepoParentPath" placeholder="Pasta-pai para o clone" class="h-10 min-w-0 rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none">
                    <button id="newGithubRepoBrowseParentBtn" class="h-10 rounded-xl border border-slate-700/30 bg-[#091c29] px-3 text-xs">📁</button>
                </div>
            </div>
            <div class="mt-5 flex justify-end gap-2">
                <button data-close-modal="githubRepoCreateModal" class="h-10 rounded-xl border border-slate-700/30 px-4 text-xs font-bold text-slate-400">Cancelar</button>
                <button id="confirmCreateGithubRepoBtn" class="h-10 rounded-xl border border-emerald-400/20 bg-emerald-400/10 px-4 text-xs font-bold text-emerald-200">Criar repositório</button>
            </div>
        </div>
    </div>

    <div id="githubRepoEditModal" class="modal-shell fixed inset-0 z-[80] hidden items-center justify-center p-4">
        <div class="panel max-h-[92vh] w-full max-w-2xl overflow-y-auto rounded-2xl p-5 shadow-2xl">
            <div class="flex items-start justify-between gap-4">
                <div><h3 class="text-lg font-bold text-slate-100">Editar repositório</h3><p id="editGithubRepoFullName" class="mt-1 text-xs text-cyan-300">—</p></div>
                <button data-close-modal="githubRepoEditModal" class="rounded-lg px-2 py-1 text-slate-500">✕</button>
            </div>
            <div class="mt-5 grid gap-3">
                <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Descrição</span><textarea id="editGithubRepoDescription" rows="3" maxlength="350" class="w-full resize-none rounded-xl border border-slate-700/30 bg-[#061722] px-3 py-3 text-sm text-slate-200 outline-none"></textarea></label>
                <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Homepage</span><input id="editGithubRepoHomepage" maxlength="500" placeholder="https://..." class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none"></label>
                <div class="grid gap-3 sm:grid-cols-2">
                    <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Visibilidade</span><select id="editGithubRepoVisibility" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-300 outline-none"><option value="private">Privado</option><option value="public">Público</option><option value="internal">Interno</option></select></label>
                    <label><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Branch padrão</span><input id="editGithubRepoDefaultBranch" class="h-11 w-full rounded-xl border border-slate-700/30 bg-[#061722] px-3 text-sm text-slate-200 outline-none"></label>
                </div>
                <div class="grid grid-cols-1 gap-2 sm:grid-cols-3">
                    <label class="rounded-xl border border-slate-700/20 bg-[#061722] p-3 text-xs text-slate-400"><input id="editGithubRepoIssues" type="checkbox" class="mr-2 accent-cyan-400"> Issues</label>
                    <label class="rounded-xl border border-slate-700/20 bg-[#061722] p-3 text-xs text-slate-400"><input id="editGithubRepoProjects" type="checkbox" class="mr-2 accent-cyan-400"> Projects</label>
                    <label class="rounded-xl border border-slate-700/20 bg-[#061722] p-3 text-xs text-slate-400"><input id="editGithubRepoWiki" type="checkbox" class="mr-2 accent-cyan-400"> Wiki</label>
                </div>
                <div class="mt-2 rounded-xl border border-red-900/25 bg-red-950/10 p-4">
                    <div class="text-[10px] font-bold uppercase tracking-[.10em] text-red-300/80">Administração avançada</div>
                    <div class="mt-3 grid gap-2 sm:grid-cols-[minmax(0,1fr)_auto]">
                        <input id="renameGithubRepoName" placeholder="novo-nome" class="h-10 min-w-0 rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none">
                        <button id="renameGithubRepoBtn" class="h-10 rounded-lg border border-amber-400/20 bg-amber-400/5 px-4 text-xs font-bold text-amber-200">Renomear</button>
                    </div>
                    <div class="mt-3 flex flex-wrap gap-2">
                        <button id="archiveGithubRepoBtn" class="h-9 rounded-lg border border-amber-400/20 bg-amber-400/5 px-3 text-[10px] font-bold text-amber-200">Arquivar</button>
                        <button id="deleteGithubRepoBtn" class="h-9 rounded-lg border border-red-900/40 bg-red-950/30 px-3 text-[10px] font-bold text-red-300">Excluir permanentemente</button>
                    </div>
                </div>
            </div>
            <div class="mt-5 flex justify-end gap-2">
                <button data-close-modal="githubRepoEditModal" class="h-10 rounded-xl border border-slate-700/30 px-4 text-xs font-bold text-slate-400">Cancelar</button>
                <button id="confirmEditGithubRepoBtn" class="h-10 rounded-xl border border-cyan-400/20 bg-cyan-400/10 px-4 text-xs font-bold text-cyan-200">Salvar alterações</button>
            </div>
        </div>
    </div>

    <div id="githubRepoCloneModal" class="modal-shell fixed inset-0 z-[80] hidden items-center justify-center p-4">
        <div class="panel w-full max-w-lg rounded-2xl p-5 shadow-2xl">
            <div class="flex items-start justify-between gap-4">
                <div><h3 class="text-lg font-bold text-slate-100">Clonar repositório</h3><p id="cloneGithubRepoFullName" class="mt-1 text-xs text-cyan-300">—</p></div>
                <button data-close-modal="githubRepoCloneModal" class="rounded-lg px-2 py-1 text-slate-500">✕</button>
            </div>
            <div class="mt-5">
                <span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Pasta-pai</span>
                <div class="grid grid-cols-[minmax(0,1fr)_auto] gap-2">
                    <input id="cloneGithubRepoParentPath" class="h-10 min-w-0 rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none">
                    <button id="cloneGithubRepoBrowseParentBtn" class="h-10 rounded-lg border border-slate-700/30 bg-[#091c29] px-3 text-xs">📁</button>
                </div>
                <label class="mt-3 block"><span class="mb-1 block text-[10px] font-bold uppercase text-slate-600">Nome da pasta</span><input id="cloneGithubRepoFolderName" class="h-10 w-full rounded-lg border border-slate-700/30 bg-[#061722] px-3 text-xs text-slate-200 outline-none"></label>
            </div>
            <div class="mt-5 flex justify-end gap-2">
                <button data-close-modal="githubRepoCloneModal" class="h-10 rounded-xl border border-slate-700/30 px-4 text-xs font-bold text-slate-400">Cancelar</button>
                <button id="confirmCloneGithubRepoBtn" class="h-10 rounded-xl border border-blue-400/20 bg-blue-400/10 px-4 text-xs font-bold text-blue-200">Clonar e selecionar</button>
            </div>
        </div>
    </div>



    <!-- MODAL CRIAR EXECUTÁVEL - GIT & GITHUB -->
    <div id="gitBuildExecutableModal" class="modal-shell fixed inset-0 z-[90] hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="panel max-h-[92vh] w-full max-w-3xl overflow-y-auto rounded-2xl shadow-2xl">
            <div class="flex items-start justify-between gap-4 border-b border-slate-700/20 px-5 py-4">
                <div>
                    <div class="text-[10px] font-bold uppercase tracking-[.12em] text-cyan-400/60">Build local</div>
                    <h3 class="mt-1 text-lg font-bold text-slate-100">Criar Executável</h3>
                    <p class="mt-1 text-xs text-slate-500">Gera TechToolHub.exe com PyInstaller e acompanha o processo em tempo real.</p>
                </div>
                <button id="gitCloseBuildExecutableModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>

            <div class="space-y-4 p-5">
                <div id="gitBuildEnvironmentBox" class="rounded-xl border border-slate-700/25 bg-[#061722] p-4 text-xs text-slate-500">
                    Verificando ambiente...
                </div>

                <div class="grid gap-3 sm:grid-cols-3">
                    <div class="status-card rounded-xl p-3">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Status</div>
                        <div id="gitBuildStatusLabel" class="mt-1 text-sm font-bold text-slate-300">Pronto</div>
                    </div>
                    <div class="status-card rounded-xl p-3 sm:col-span-2">
                        <div class="text-[9px] font-bold uppercase tracking-[.12em] text-slate-600">Saída</div>
                        <div id="gitBuildOutputPath" class="mt-1 truncate text-xs font-semibold text-slate-400">—</div>
                    </div>
                </div>

                <div>
                    <div class="mb-2 flex items-center justify-between gap-3">
                        <span class="text-xs font-bold text-slate-300">Log do build</span>
                        <span id="gitBuildRunningBadge" class="hidden rounded-md border border-cyan-400/15 bg-cyan-400/5 px-2 py-1 text-[9px] font-bold uppercase tracking-[.08em] text-cyan-300">
                            Processando
                        </span>
                    </div>
                    <pre id="gitBuildLog" class="log-box max-h-[320px] min-h-[170px] overflow-auto whitespace-pre-wrap rounded-xl p-4 text-[11px] leading-5 text-slate-500">Aguardando início do build...</pre>
                </div>

                <div id="gitBuildError" class="hidden rounded-lg border border-red-900/50 bg-red-950/30 px-4 py-3 text-xs text-red-300"></div>

                <div class="flex flex-col-reverse gap-2 sm:flex-row sm:justify-between">
                    <button id="gitOpenBuildFolderBtn" type="button" disabled class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400 disabled:cursor-not-allowed disabled:opacity-35">
                        Abrir pasta de saída
                    </button>
                    <div class="flex flex-col-reverse gap-2 sm:flex-row">
                        <button id="gitCancelBuildModalBtn" type="button" class="h-10 rounded-lg border border-slate-700/35 px-4 text-xs font-semibold text-slate-400">
                            Fechar
                        </button>
                        <button id="gitRunExecutableBuildBtn" type="button" class="h-10 rounded-lg border border-cyan-300/20 bg-cyan-400/10 px-5 text-xs font-bold text-cyan-200 disabled:cursor-not-allowed disabled:opacity-40">
                            Gerar EXE
                        </button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- MODAL README - GIT & GITHUB -->
    <div id="gitReadmeModal" class="modal-shell fixed inset-0 z-[90] hidden items-center justify-center p-4" role="dialog" aria-modal="true">
        <div class="panel max-h-[92vh] w-full max-w-4xl overflow-y-auto rounded-2xl shadow-2xl">
            <div class="flex items-start justify-between gap-4 border-b border-slate-700/20 px-5 py-4">
                <div>
                    <div class="text-[10px] font-bold uppercase tracking-[.12em] text-cyan-400/60">Documentação</div>
                    <h3 class="mt-1 text-lg font-bold text-slate-100">README — TECH TOOL HUB</h3>
                    <p class="mt-1 text-xs text-slate-500">Execução, dados, build do EXE e empacotamento MSIX.</p>
                </div>
                <button id="gitCloseReadmeModalBtn" type="button" class="rounded-lg p-2 text-slate-500 hover:bg-slate-800/40 hover:text-slate-200">✕</button>
            </div>
            <div class="p-5">
                <pre id="gitReadmeContent" class="log-box max-h-[70vh] overflow-auto whitespace-pre-wrap rounded-xl p-4 text-xs leading-6 text-slate-400">Carregando README...</pre>
            </div>
        </div>
    </div>

    <div id="gitToast" class="pointer-events-none fixed bottom-5 right-5 z-[70] hidden max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl"></div>

    <script>
        const state = {
            environment: null,
            audit: null,
            githubRepos: [],
            githubRepoSummary: {},
            githubLogin: '',
            editingGithubRepo: null,
            cloningGithubRepo: null
        };

        const environmentMessage = document.getElementById('environmentMessage');
        const sourceModeBadge = document.getElementById('sourceModeBadge');
        const projectSelect = document.getElementById('projectSelect');
        const selectedProjectPath = document.getElementById('selectedProjectPath');
        const selectedProjectGitBadge = document.getElementById('selectedProjectGitBadge');
        const gitBranch = document.getElementById('gitBranch');
        const gitChangedFiles = document.getElementById('gitChangedFiles');
        const ghStatus = document.getElementById('ghStatus');
        const latestReleaseName = document.getElementById('latestReleaseName');
        const gitRemote = document.getElementById('gitRemote');

        const auditCandidates = document.getElementById('auditCandidates');
        const auditBlocking = document.getElementById('auditBlocking');
        const auditWarnings = document.getElementById('auditWarnings');
        const auditResult = document.getElementById('auditResult');

        const commitMessageInput = document.getElementById('commitMessageInput');
        const publishResult = document.getElementById('publishResult');

        const releaseTagInput = document.getElementById('releaseTagInput');
        const releaseTitleInput = document.getElementById('releaseTitleInput');
        const releaseAssetSelect = document.getElementById('releaseAssetSelect');
        const releaseNotesInput = document.getElementById('releaseNotesInput');
        const includeCertificateInput = document.getElementById('includeCertificateInput');
        const releaseDraftInput = document.getElementById('releaseDraftInput');
        const releaseResult = document.getElementById('releaseResult');

        const gitActionLog = document.getElementById('gitActionLog');
        const gitToast = document.getElementById('gitToast');


        const gitCreateExecutableBtn = document.getElementById('gitCreateExecutableBtn');
        const gitShowReadmeBtn = document.getElementById('gitShowReadmeBtn');

        const gitBuildExecutableModal = document.getElementById('gitBuildExecutableModal');
        const gitBuildEnvironmentBox = document.getElementById('gitBuildEnvironmentBox');
        const gitBuildStatusLabel = document.getElementById('gitBuildStatusLabel');
        const gitBuildOutputPath = document.getElementById('gitBuildOutputPath');
        const gitBuildLog = document.getElementById('gitBuildLog');
        const gitBuildRunningBadge = document.getElementById('gitBuildRunningBadge');
        const gitBuildError = document.getElementById('gitBuildError');
        const gitRunExecutableBuildBtn = document.getElementById('gitRunExecutableBuildBtn');
        const gitOpenBuildFolderBtn = document.getElementById('gitOpenBuildFolderBtn');

        const gitReadmeModal = document.getElementById('gitReadmeModal');
        const gitReadmeContent = document.getElementById('gitReadmeContent');

        let gitBuildPollTimer = null;


        const githubRepoGrid = document.getElementById('githubRepoGrid');
        const githubReposLoading = document.getElementById('githubReposLoading');
        const githubReposError = document.getElementById('githubReposError');
        const githubRepoSearchInput = document.getElementById('githubRepoSearchInput');
        const githubRepoVisibilityFilter = document.getElementById('githubRepoVisibilityFilter');
        const githubRepoStateFilter = document.getElementById('githubRepoStateFilter');
        const githubRepoSortSelect = document.getElementById('githubRepoSortSelect');
        const githubRepoScopeSelect = document.getElementById('githubRepoScopeSelect');

        function escapeHtml(value) {
            return String(value ?? '')
                .replaceAll('&', '&amp;')
                .replaceAll('<', '&lt;')
                .replaceAll('>', '&gt;')
                .replaceAll('"', '&quot;')
                .replaceAll("'", '&#039;');
        }

        function basename(path) {
            return String(path || '').split(/[\\/]/).pop() || '';
        }

        function showToast(message, error = false) {
            gitToast.textContent = message;
            gitToast.className =
                'pointer-events-none fixed bottom-5 right-5 z-[70] max-w-sm rounded-xl border px-4 py-3 text-sm shadow-2xl ' +
                (error
                    ? 'border-red-900/60 bg-red-950/95 text-red-200'
                    : 'border-cyan-900/60 bg-[#092633]/95 text-cyan-100');

            gitToast.classList.remove('hidden');
            clearTimeout(showToast.timer);
            showToast.timer = setTimeout(() => gitToast.classList.add('hidden'), 2600);
        }

        function appendLog(message) {
            const stamp = new Date().toLocaleTimeString('pt-BR');
            const current = gitActionLog.textContent === 'Pronto.' ? '' : gitActionLog.textContent;
            gitActionLog.textContent = `${current}${current ? '\n\n' : ''}[${stamp}] ${message}`;
            gitActionLog.scrollTop = gitActionLog.scrollHeight;
        }


        function setGitUtilityModalVisible(modal, visible) {
            if (!modal) return;
            modal.classList.toggle('hidden', !visible);
            modal.classList.toggle('flex', visible);
        }

        function gitBuildStatusText(status) {
            if (status === 'starting') return 'Preparando...';
            if (status === 'running') return 'Gerando...';
            if (status === 'success') return 'Concluído';
            if (status === 'error') return 'Falhou';
            return 'Pronto';
        }

        function stopGitBuildPolling() {
            if (!gitBuildPollTimer) return;
            clearInterval(gitBuildPollTimer);
            gitBuildPollTimer = null;
        }

        function startGitBuildPolling() {
            if (gitBuildPollTimer) return;
            gitBuildPollTimer = setInterval(refreshGitBuildStatus, 1000);
        }

        function renderGitBuildState(data) {
            const environment = data.environment || {};
            const reasons = Array.isArray(environment.reasons)
                ? environment.reasons
                : [];

            gitBuildEnvironmentBox.innerHTML = environment.available
                ? `<div class="flex items-start gap-3">
                       <span class="mt-1 h-2 w-2 shrink-0 rounded-full bg-emerald-400"></span>
                       <div>
                           <div class="font-semibold text-emerald-300">Ambiente pronto para gerar o EXE.</div>
                           <div class="mt-1 text-[10px] text-slate-600">Build local do TECH TOOL HUB.</div>
                       </div>
                   </div>`
                : `<div class="flex items-start gap-3">
                       <span class="mt-1 h-2 w-2 shrink-0 rounded-full bg-amber-300"></span>
                       <div>
                           <div class="font-semibold text-amber-200">Build indisponível neste modo.</div>
                           <div class="mt-1 space-y-1 text-[10px] text-slate-600">
                               ${reasons.map(reason => `<div>• ${escapeHtml(reason)}</div>`).join('')}
                           </div>
                       </div>
                   </div>`;

            gitBuildStatusLabel.textContent = gitBuildStatusText(data.status);
            gitBuildOutputPath.textContent =
                data.output_path
                || environment.output_dir
                || '—';

            gitBuildLog.textContent =
                (data.log || []).join('\n')
                || data.message
                || 'Aguardando início do build...';

            const running = Boolean(data.running);
            gitBuildRunningBadge.classList.toggle('hidden', !running);
            gitRunExecutableBuildBtn.disabled =
                running || !environment.available;
            gitRunExecutableBuildBtn.textContent =
                running ? 'Gerando...' : 'Gerar EXE';
            gitOpenBuildFolderBtn.disabled =
                data.status !== 'success';

            if (data.status === 'error') {
                gitBuildError.textContent =
                    data.message || 'O build falhou.';
                gitBuildError.classList.remove('hidden');
            } else {
                gitBuildError.classList.add('hidden');
                gitBuildError.textContent = '';
            }

            if (running) startGitBuildPolling();
            else stopGitBuildPolling();
        }

        async function refreshGitBuildStatus() {
            try {
                const response = await fetch(
                    '/api/build/status',
                    { cache: 'no-store' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(
                        data.detail
                        || 'Não foi possível consultar o build.'
                    );
                }

                renderGitBuildState(data);
            } catch (error) {
                gitBuildError.textContent =
                    error.message
                    || 'Erro ao consultar o build.';
                gitBuildError.classList.remove('hidden');
            }
        }

        async function openGitBuildExecutableModal() {
            setGitUtilityModalVisible(
                gitBuildExecutableModal,
                true
            );
            await refreshGitBuildStatus();
        }

        function closeGitBuildExecutableModal() {
            setGitUtilityModalVisible(
                gitBuildExecutableModal,
                false
            );
            stopGitBuildPolling();
        }

        async function runGitExecutableBuild() {
            gitRunExecutableBuildBtn.disabled = true;
            gitBuildError.classList.add('hidden');

            try {
                const response = await fetch(
                    '/api/build/executable',
                    { method: 'POST' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(
                        data.detail
                        || 'Não foi possível iniciar o build.'
                    );
                }

                renderGitBuildState(data);
                startGitBuildPolling();
                appendLog('Build do executável iniciado.');
            } catch (error) {
                gitBuildError.textContent =
                    error.message
                    || 'Erro ao iniciar o build.';
                gitBuildError.classList.remove('hidden');
                await refreshGitBuildStatus();
            }
        }

        async function openGitBuildOutputFolder() {
            try {
                const response = await fetch(
                    '/api/build/open-output',
                    { method: 'POST' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(
                        data.detail
                        || 'Não foi possível abrir a pasta.'
                    );
                }

                showToast(data.message || 'Pasta aberta.');
            } catch (error) {
                showToast(
                    error.message || 'Erro ao abrir a pasta.',
                    true
                );
            }
        }

        async function openGitReadmeModal() {
            setGitUtilityModalVisible(gitReadmeModal, true);
            gitReadmeContent.textContent =
                'Carregando README...';

            try {
                const response = await fetch(
                    '/api/readme',
                    { cache: 'no-store' }
                );
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(
                        data.detail
                        || 'Não foi possível carregar o README.'
                    );
                }

                gitReadmeContent.textContent =
                    data.content || 'README vazio.';
            } catch (error) {
                gitReadmeContent.textContent =
                    `Erro ao carregar README:\n${error.message || error}`;
            }
        }

        function closeGitReadmeModal() {
            setGitUtilityModalVisible(gitReadmeModal, false);
        }

        function resultBox(element, message, error = false, link = '') {
            element.className =
                'mt-4 rounded-xl border px-3 py-3 text-xs ' +
                (error
                    ? 'border-red-900/40 bg-red-950/20 text-red-300'
                    : 'border-emerald-900/35 bg-emerald-950/15 text-emerald-300');

            element.innerHTML = `
                <div>${escapeHtml(message)}</div>
                ${link ? `<a href="${escapeHtml(link)}" target="_blank" rel="noopener noreferrer" class="mt-2 inline-block font-bold text-cyan-300 hover:text-cyan-200">Abrir Release ↗</a>` : ''}
            `;
            element.classList.remove('hidden');
        }

        function renderProjectSelector(data) {
            const projects = Array.isArray(data.projects) ? data.projects : [];
            const selected = data.selected_project || {};
            const selectedPath = selected.path || '';

            projectSelect.innerHTML = '';

            if (!projects.length) {
                const option = document.createElement('option');
                option.value = '';
                option.textContent = 'Nenhum projeto salvo';
                projectSelect.appendChild(option);
            } else {
                projects.forEach(project => {
                    const option = document.createElement('option');
                    option.value = project.path;
                    option.textContent = `${project.name}${project.is_git ? '' : ' — sem Git'}`;
                    option.selected = project.path === selectedPath;
                    projectSelect.appendChild(option);
                });
            }

            selectedProjectPath.textContent = selectedPath || 'Nenhuma pasta selecionada';
            selectedProjectGitBadge.textContent = selected.is_git ? 'Git pronto' : 'Sem Git';
            selectedProjectGitBadge.className =
                'rounded-md border px-2 py-0.5 text-[9px] font-bold uppercase tracking-[.08em] ' +
                (selected.is_git
                    ? 'border-emerald-400/20 bg-emerald-400/5 text-emerald-300'
                    : 'border-amber-400/20 bg-amber-400/5 text-amber-200');
        }

        function renderReleaseAssets(assets) {
            const files = Array.isArray(assets?.files)
                ? assets.files.filter(item => !item.is_certificate)
                : [];

            releaseAssetSelect.innerHTML =
                '<option value="">Sem artefato / detectar automaticamente</option>';

            files.forEach(item => {
                const option = document.createElement('option');
                option.value = item.path;

                const relativePath = item.relative_path || item.name;
                const isExe = String(item.name || '').toLowerCase().endsWith('.exe');

                option.textContent = isExe
                    ? `EXE • ${relativePath}`
                    : relativePath;

                option.selected = Boolean(
                    assets.primary
                    && item.path === assets.primary
                );

                releaseAssetSelect.appendChild(option);
            });

            latestReleaseName.textContent = assets?.primary
                ? basename(assets.primary)
                : 'Nenhum artefato';
        }

        async function chooseProjectFolder() {
            const button = document.getElementById('chooseProjectFolderBtn');
            button.disabled = true;
            button.textContent = 'Abrindo seletor...';

            try {
                const response = await fetch('/api/git-dashboard/select-folder', {
                    method: 'POST'
                });
                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha ao selecionar pasta.');

                if (data.cancelled) {
                    appendLog('Seleção de pasta cancelada.');
                    return;
                }

                appendLog(
                    `Projeto selecionado: ${data.selected_path}` +
                    (data.is_git ? '' : '\nA pasta ainda não é um repositório Git.')
                );
                showToast(data.is_git ? 'Projeto Git selecionado.' : 'Pasta selecionada; Git ainda não inicializado.');
                await loadEnvironment({ resetReleaseFields: true });
            } catch (error) {
                showToast(error.message || 'Erro ao selecionar pasta.', true);
                appendLog(`ERRO seletor: ${error.message || error}`);
            } finally {
                button.disabled = false;
                button.textContent = '📁 Selecionar pasta';
            }
        }

        async function activateProject(path) {
            if (!path) return;

            try {
                const response = await fetch('/api/git-dashboard/activate-project', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path })
                });
                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha ao ativar projeto.');

                appendLog(`Projeto ativo: ${data.selected_path}`);
                await loadEnvironment({ resetReleaseFields: true });
            } catch (error) {
                showToast(error.message || 'Erro ao trocar projeto.', true);
            }
        }

        async function forgetActiveProject() {
            const path = projectSelect.value;
            if (!path) return;

            try {
                const response = await fetch('/api/git-dashboard/forget-project', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ path })
                });
                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha ao remover projeto da lista.');

                appendLog(`Projeto removido da lista: ${path}`);
                await loadEnvironment({ resetReleaseFields: true });
            } catch (error) {
                showToast(error.message || 'Erro ao remover projeto.', true);
            }
        }

        async function loadEnvironment(options = {}) {
            const refreshButton = document.getElementById('refreshGitStatusBtn');
            refreshButton.disabled = true;

            try {
                const response = await fetch('/api/git-dashboard/status', { cache: 'no-store' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Falha ao consultar ambiente.');

                state.environment = data;
                renderProjectSelector(data);

                gitBranch.textContent = data.branch || '—';
                gitChangedFiles.textContent = data.changed_files ?? 0;
                gitRemote.textContent = data.remote || 'origin não configurado';

                sourceModeBadge.textContent = data.source_mode ? 'Projeto pronto' : 'Indisponível';
                sourceModeBadge.className =
                    'rounded-lg border px-2.5 py-1 text-[10px] font-bold uppercase tracking-[.08em] ' +
                    (data.source_mode
                        ? 'border-emerald-400/20 bg-emerald-400/5 text-emerald-300'
                        : 'border-amber-400/20 bg-amber-400/5 text-amber-200');

                environmentMessage.textContent = data.source_mode
                    ? `Repositório: ${data.repo_root}${data.behind ? ` • remoto +${data.behind}` : ''}${data.ahead ? ` • local +${data.ahead}` : ''}`
                    : (data.source_reason || 'Git indisponível para o projeto selecionado.');

                if (data.gh_authenticated) {
                    ghStatus.textContent = data.gh_detection_source
                        ? `Autenticado • ${data.gh_detection_source}`
                        : 'Autenticado';
                } else if (data.gh_installed) {
                    ghStatus.textContent = data.gh_detection_source
                        ? `Login necessário • ${data.gh_detection_source}`
                        : 'Login necessário';
                } else {
                    ghStatus.textContent = 'Não localizado';
                }

                const assets = data.release_assets || {};
                renderReleaseAssets(assets);

                if (options.resetReleaseFields) {
                    releaseTagInput.value = assets.suggested_tag || '';
                    releaseTitleInput.value =
                        assets.suggested_title ||
                        (data.selected_project?.name
                            ? `${data.selected_project.name} Release`
                            : '');
                    releaseNotesInput.value = '';
                } else {
                    if (!releaseTagInput.value && assets.suggested_tag) {
                        releaseTagInput.value = assets.suggested_tag;
                    }

                    if (!releaseTitleInput.value) {
                        releaseTitleInput.value =
                            assets.suggested_title ||
                            (data.selected_project?.name
                                ? `${data.selected_project.name} Release`
                                : '');
                    }
                }

                const actionsAvailable = Boolean(data.source_mode);
                const hasSelectedFolder = Boolean(data.selected_project?.path);
                const isGitRepo = Boolean(data.selected_project?.is_git);

                document.getElementById('runSecurityAuditBtn').disabled = !actionsAvailable;
                document.getElementById('publishGitBtn').disabled = !actionsAvailable;
                document.getElementById('createReleaseBtn').disabled =
                    !actionsAvailable ||
                    !data.gh_authenticated;

                document.getElementById('gitInitBtn').disabled =
                    !data.git_installed ||
                    !hasSelectedFolder ||
                    isGitRepo;

                document.querySelectorAll('.git-repo-command').forEach(button => {
                    button.disabled = !isGitRepo;
                    button.classList.toggle('opacity-40', !isGitRepo);
                    button.classList.toggle('cursor-not-allowed', !isGitRepo);
                });

                [
                    'gitAddBtn',
                    'gitCommitBtn',
                    'gitSwitchBtn',
                    'gitSwitchCreateBtn',
                    'gitBranchCreateBtn',
                    'gitBranchDeleteBtn',
                    'gitRmCachedBtn'
                ].forEach(id => {
                    const button = document.getElementById(id);
                    button.disabled = !isGitRepo;
                    button.classList.toggle('opacity-40', !isGitRepo);
                    button.classList.toggle('cursor-not-allowed', !isGitRepo);
                });

                const commandBadge = document.getElementById('gitIndividualRepoBadge');
                commandBadge.textContent = isGitRepo
                    ? `${data.selected_project.name || 'Projeto'} • ${data.branch || 'HEAD'}`
                    : (hasSelectedFolder ? 'pasta sem Git' : 'sem projeto');

                commandBadge.className =
                    'rounded-lg border bg-[#061722] px-3 py-1.5 text-[10px] font-bold uppercase tracking-[.08em] ' +
                    (isGitRepo
                        ? 'border-emerald-400/20 text-emerald-300'
                        : 'border-amber-400/20 text-amber-200');

                const cloneParent = document.getElementById('gitCloneParentInput');
                if (!cloneParent.value && data.selected_project?.path) {
                    const currentPath = data.selected_project.path;
                    const separator = currentPath.includes('\\') ? '\\' : '/';
                    const parts = currentPath.split(/[\\/]/).filter(Boolean);

                    if (parts.length > 1) {
                        const isWindowsDrive = /^[A-Za-z]:/.test(currentPath);
                        if (isWindowsDrive) {
                            cloneParent.value = parts.slice(0, -1).join('\\');
                        } else {
                            cloneParent.value = '/' + parts.slice(0, -1).join('/');
                        }
                    }
                }

            } catch (error) {
                showToast(error.message || 'Erro ao carregar status.', true);
            } finally {
                refreshButton.disabled = false;
            }
        }

        async function runSecurityAudit() {
            const button = document.getElementById('runSecurityAuditBtn');
            button.disabled = true;
            button.textContent = 'Verificando...';
            auditResult.classList.add('hidden');

            try {
                const response = await fetch('/api/git-dashboard/security-check', {
                    method: 'POST'
                });
                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha na auditoria.');

                state.audit = data;

                auditCandidates.textContent = data.candidate_count ?? 0;
                auditBlocking.textContent = data.blocking_count ?? 0;
                auditWarnings.textContent = data.warning_count ?? 0;

                const findings = Array.isArray(data.findings) ? data.findings : [];

                auditResult.className = 'mt-4 space-y-2';
                auditResult.innerHTML = `
                    <div class="rounded-xl border px-3 py-3 text-xs ${
                        data.safe
                            ? 'border-emerald-900/35 bg-emerald-950/15 text-emerald-300'
                            : 'border-red-900/40 bg-red-950/20 text-red-300'
                    }">
                        ${data.safe
                            ? '✓ Nenhum risco alto encontrado.'
                            : `⚠ ${data.blocking_count} risco(s) alto(s) bloqueiam a publicação.`}
                    </div>

                    ${findings.slice(0, 12).map(finding => `
                        <div class="rounded-lg border p-2.5 ${
                            finding.severity === 'high'
                                ? 'finding-high'
                                : 'finding-warning'
                        }">
                            <div class="text-[10px] font-bold ${
                                finding.severity === 'high'
                                    ? 'text-red-300'
                                    : 'text-amber-200'
                            }">${escapeHtml(finding.file)}${finding.line ? `:${finding.line}` : ''}</div>
                            <div class="mt-1 text-[10px] leading-4 text-slate-500">${escapeHtml(finding.message)}</div>
                        </div>
                    `).join('')}
                `;

                appendLog(
                    `Auditoria: ${data.blocking_count} risco(s) alto(s), ` +
                    `${data.warning_count} aviso(s), ${data.candidate_count} arquivo(s) analisados.`
                );

            } catch (error) {
                resultBox(auditResult, error.message || 'Erro na auditoria.', true);
                appendLog(`ERRO auditoria: ${error.message || error}`);
            } finally {
                button.disabled = false;
                button.textContent = 'Executar checagem';
            }
        }

        async function publishProject() {
            const button = document.getElementById('publishGitBtn');
            const commitMessage = commitMessageInput.value.trim();

            if (!commitMessage) {
                showToast('Informe a mensagem do commit.', true);
                commitMessageInput.focus();
                return;
            }

            button.disabled = true;
            button.textContent = 'Publicando...';
            publishResult.classList.add('hidden');

            try {
                const response = await fetch('/api/git-dashboard/publish', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ commit_message: commitMessage })
                });

                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha na publicação.');

                resultBox(
                    publishResult,
                    data.committed
                        ? `Commit criado e branch ${data.branch} enviada ao GitHub.`
                        : `Nenhuma mudança nova para commit. Branch ${data.branch} sincronizada.`
                );

                appendLog((data.log || []).join('\n'));
                showToast('Projeto publicado no GitHub.');
                await loadEnvironment();

            } catch (error) {
                resultBox(publishResult, error.message || 'Erro ao publicar.', true);
                appendLog(`ERRO publicação: ${error.message || error}`);
            } finally {
                button.disabled = false;
                button.textContent = 'Auditar e publicar';
            }
        }

        async function createRelease() {
            const button = document.getElementById('createReleaseBtn');

            const payload = {
                tag: releaseTagInput.value.trim(),
                title: releaseTitleInput.value.trim(),
                notes: releaseNotesInput.value.trim(),
                asset_path: releaseAssetSelect.value,
                include_certificate: includeCertificateInput.checked,
                draft: releaseDraftInput.checked
            };

            if (!payload.tag || !payload.title) {
                showToast('Informe a tag e o título da Release.', true);
                return;
            }

            button.disabled = true;
            button.textContent = 'Criando Release...';
            releaseResult.classList.add('hidden');

            try {
                const response = await fetch('/api/git-dashboard/release', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const data = await response.json();

                if (!response.ok) throw new Error(data.detail || 'Falha ao criar Release.');

                resultBox(
                    releaseResult,
                    `${data.tag} criada com sucesso.`,
                    false,
                    data.url || ''
                );

                appendLog(
                    `Release ${data.tag}\nMSIX: ${data.msix}` +
                    (data.certificate ? `\nCertificado: ${data.certificate}` : '') +
                    (data.url ? `\nURL: ${data.url}` : '')
                );

                showToast('GitHub Release criada.');
                await loadEnvironment();

            } catch (error) {
                resultBox(releaseResult, error.message || 'Erro ao criar Release.', true);
                appendLog(`ERRO Release: ${error.message || error}`);
            } finally {
                button.disabled = false;
                button.textContent = 'Criar GitHub Release';
            }
        }


        const gitIndividualResult = document.getElementById('gitIndividualResult');


        function formatRepoDate(value) {
            if (!value) return '—';
            const date = new Date(value);
            if (Number.isNaN(date.getTime())) return value;
            return date.toLocaleDateString('pt-BR');
        }

        function formatRepoSize(sizeKb) {
            const kb = Number(sizeKb || 0);
            if (kb < 1024) return `${kb} KB`;
            const mb = kb / 1024;
            if (mb < 1024) return `${mb.toFixed(mb >= 100 ? 0 : 1)} MB`;
            return `${(mb / 1024).toFixed(2)} GB`;
        }

        function openModal(id) {
            const modal = document.getElementById(id);
            modal.classList.remove('hidden');
            modal.classList.add('flex');
        }

        function closeModal(id) {
            const modal = document.getElementById(id);
            modal.classList.add('hidden');
            modal.classList.remove('flex');
        }

        async function browseFolderForInput(inputId) {
            try {
                const response = await fetch('/api/git-dashboard/browse-folder', { method: 'POST' });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Falha ao selecionar pasta.');
                if (data.path) document.getElementById(inputId).value = data.path;
            } catch (error) {
                showToast(error.message || 'Erro ao selecionar pasta.', true);
            }
        }

        function repoVisibilityBadge(repo) {
            if (repo.visibility === 'private') return '<span class="rounded-md border border-violet-400/20 bg-violet-400/5 px-2 py-0.5 text-[9px] font-bold uppercase text-violet-200">Privado</span>';
            if (repo.visibility === 'internal') return '<span class="rounded-md border border-amber-400/20 bg-amber-400/5 px-2 py-0.5 text-[9px] font-bold uppercase text-amber-200">Interno</span>';
            return '<span class="rounded-md border border-cyan-400/20 bg-cyan-400/5 px-2 py-0.5 text-[9px] font-bold uppercase text-cyan-200">Público</span>';
        }

        function getFilteredGithubRepos() {
            const query = githubRepoSearchInput.value.trim().toLowerCase();
            const visibility = githubRepoVisibilityFilter.value;
            const stateFilter = githubRepoStateFilter.value;
            const sort = githubRepoSortSelect.value;
            let repos = [...state.githubRepos];

            if (query) {
                repos = repos.filter(repo => [
                    repo.full_name, repo.description, repo.language, repo.default_branch
                ].join(' ').toLowerCase().includes(query));
            }

            if (visibility !== 'all') repos = repos.filter(repo => repo.visibility === visibility);
            if (stateFilter === 'active') repos = repos.filter(repo => !repo.archived);
            else if (stateFilter === 'archived') repos = repos.filter(repo => repo.archived);
            else if (stateFilter === 'local') repos = repos.filter(repo => Boolean(repo.local));
            else if (stateFilter === 'fork') repos = repos.filter(repo => repo.fork);

            if (sort === 'name') repos.sort((a, b) => a.full_name.localeCompare(b.full_name, 'pt-BR'));
            else if (sort === 'stars') repos.sort((a, b) => Number(b.stars || 0) - Number(a.stars || 0));
            else if (sort === 'size') repos.sort((a, b) => Number(b.size_kb || 0) - Number(a.size_kb || 0));
            else repos.sort((a, b) => String(b.updated_at || '').localeCompare(String(a.updated_at || '')));

            return repos;
        }

        function renderGithubRepos() {
            const repos = getFilteredGithubRepos();
            document.getElementById('githubRepoFilteredCount').textContent =
                `${repos.length} exibido(s) de ${state.githubRepos.length}`;

            if (!repos.length) {
                githubRepoGrid.innerHTML = '<div class="md:col-span-2 2xl:col-span-3 rounded-xl border border-slate-700/20 bg-[#061722] p-6 text-center text-xs text-slate-500">Nenhum repositório corresponde aos filtros.</div>';
                return;
            }

            githubRepoGrid.innerHTML = repos.map(repo => {
                const local = repo.local || null;
                const canAdmin = Boolean(repo.can_admin);

                return `
                    <article class="repo-card min-w-0 rounded-2xl p-4">
                        <div class="flex items-start justify-between gap-3">
                            <div class="min-w-0">
                                <div class="flex flex-wrap items-center gap-2">
                                    ${repoVisibilityBadge(repo)}
                                    ${repo.archived ? '<span class="rounded-md border border-amber-400/20 px-2 py-0.5 text-[9px] font-bold uppercase text-amber-200">Arquivado</span>' : ''}
                                    ${repo.fork ? '<span class="rounded-md border border-slate-600/30 px-2 py-0.5 text-[9px] font-bold uppercase text-slate-400">Fork</span>' : ''}
                                    ${local ? '<span class="rounded-md border border-emerald-400/20 bg-emerald-400/5 px-2 py-0.5 text-[9px] font-bold uppercase text-emerald-200">No PC</span>' : ''}
                                </div>
                                <h4 class="mt-2 truncate text-sm font-bold text-slate-100">${escapeHtml(repo.name)}</h4>
                                <div class="mt-0.5 truncate text-[10px] text-slate-600">${escapeHtml(repo.full_name)}</div>
                            </div>
                            <a href="${escapeHtml(repo.url)}" target="_blank" rel="noopener noreferrer" class="shrink-0 rounded-lg border border-slate-700/25 px-2.5 py-1.5 text-[10px] font-bold text-slate-400 hover:text-cyan-200">GitHub ↗</a>
                        </div>

                        <p class="mt-3 min-h-[40px] text-xs leading-5 text-slate-500">${escapeHtml(repo.description || 'Sem descrição.')}</p>

                        <div class="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-[10px] text-slate-600">
                            <span>${repo.language ? `● ${escapeHtml(repo.language)}` : '● —'}</span>
                            <span>★ ${Number(repo.stars || 0)}</span>
                            <span>⑂ ${Number(repo.forks || 0)}</span>
                            <span>Issues ${Number(repo.open_issues || 0)}</span>
                            <span>${formatRepoSize(repo.size_kb)}</span>
                        </div>

                        <div class="mt-3 border-t border-slate-700/15 pt-3 text-[10px] text-slate-600">
                            <div class="flex justify-between gap-3">
                                <span>Branch: <strong class="font-semibold text-slate-400">${escapeHtml(repo.default_branch || '—')}</strong></span>
                                <span>${escapeHtml(formatRepoDate(repo.updated_at))}</span>
                            </div>
                            ${local ? `<div class="mt-1 truncate" title="${escapeHtml(local.path)}">Local: <span class="text-emerald-300/75">${escapeHtml(local.path)}</span>${local.changed_files ? ` • ${local.changed_files} alteração(ões)` : ''}</div>` : ''}
                        </div>

                        <div class="mt-4 grid grid-cols-2 gap-2">
                            ${local
                                ? `<button data-repo-action="select_local" data-repo="${escapeHtml(repo.full_name)}" class="h-9 rounded-lg border border-emerald-400/20 bg-emerald-400/5 px-3 text-[10px] font-bold text-emerald-200">Usar projeto local</button>`
                                : `<button data-repo-action="clone" data-repo="${escapeHtml(repo.full_name)}" class="h-9 rounded-lg border border-blue-400/20 bg-blue-400/5 px-3 text-[10px] font-bold text-blue-200">Clonar</button>`
                            }
                            <button data-repo-action="edit" data-repo="${escapeHtml(repo.full_name)}" class="h-9 rounded-lg border border-cyan-400/20 bg-cyan-400/5 px-3 text-[10px] font-bold text-cyan-200 ${canAdmin ? '' : 'opacity-40 cursor-not-allowed'}" ${canAdmin ? '' : 'disabled'}>Editar</button>
                            <button data-repo-action="${repo.archived ? 'unarchive' : 'archive'}" data-repo="${escapeHtml(repo.full_name)}" class="h-9 rounded-lg border border-amber-400/20 bg-amber-400/5 px-3 text-[10px] font-bold text-amber-200 ${canAdmin ? '' : 'opacity-40 cursor-not-allowed'}" ${canAdmin ? '' : 'disabled'}>${repo.archived ? 'Restaurar' : 'Arquivar'}</button>
                            <button data-repo-action="copy" data-repo="${escapeHtml(repo.full_name)}" class="h-9 rounded-lg border border-slate-700/25 bg-[#071c29] px-3 text-[10px] font-bold text-slate-400">Copiar URL</button>
                        </div>
                    </article>
                `;
            }).join('');

            githubRepoGrid.querySelectorAll('[data-repo-action]').forEach(button => {
                button.addEventListener('click', () => handleGithubRepoCardAction(
                    button.dataset.repoAction,
                    button.dataset.repo
                ));
            });
        }

        async function loadGithubRepositories() {
            githubReposLoading.classList.remove('hidden');
            githubReposError.classList.add('hidden');
            document.getElementById('refreshGithubReposBtn').disabled = true;

            try {
                const response = await fetch(
                    `/api/github/repos?scope=${encodeURIComponent(githubRepoScopeSelect.value || 'owner')}`,
                    { cache: 'no-store' }
                );
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Falha ao consultar repositórios.');

                state.githubRepos = Array.isArray(data.repositories) ? data.repositories : [];
                state.githubRepoSummary = data.summary || {};
                state.githubLogin = data.login || '';

                document.getElementById('githubRepoAccountMessage').textContent =
                    data.login ? `Conta: @${data.login} • consulta via GitHub CLI` : 'Conta GitHub autenticada';
                document.getElementById('repoCountTotal').textContent = data.summary?.total ?? 0;
                document.getElementById('repoCountPublic').textContent = data.summary?.public ?? 0;
                document.getElementById('repoCountPrivate').textContent = data.summary?.private ?? 0;
                document.getElementById('repoCountLocal').textContent = data.summary?.local ?? 0;
                document.getElementById('repoCountArchived').textContent = data.summary?.archived ?? 0;
                document.getElementById('repoCountForks').textContent = data.summary?.forks ?? 0;

                renderGithubRepos();
            } catch (error) {
                state.githubRepos = [];
                githubRepoGrid.innerHTML = '';
                githubReposError.textContent = error.message || String(error);
                githubReposError.classList.remove('hidden');
                document.getElementById('githubRepoAccountMessage').textContent =
                    'GitHub indisponível. Verifique gh --version e gh auth status.';
            } finally {
                githubReposLoading.classList.add('hidden');
                document.getElementById('refreshGithubReposBtn').disabled = false;
            }
        }

        function findGithubRepo(fullName) {
            return state.githubRepos.find(repo => repo.full_name === fullName) || null;
        }

        function openGithubRepoEdit(repo) {
            state.editingGithubRepo = repo;
            document.getElementById('editGithubRepoFullName').textContent = repo.full_name;
            document.getElementById('editGithubRepoDescription').value = repo.description || '';
            document.getElementById('editGithubRepoHomepage').value = repo.homepage || '';
            document.getElementById('editGithubRepoVisibility').value = repo.visibility || 'private';
            document.getElementById('editGithubRepoDefaultBranch').value = repo.default_branch || '';
            document.getElementById('editGithubRepoIssues').checked = Boolean(repo.has_issues);
            document.getElementById('editGithubRepoProjects').checked = Boolean(repo.has_projects);
            document.getElementById('editGithubRepoWiki').checked = Boolean(repo.has_wiki);
            document.getElementById('renameGithubRepoName').value = repo.name || '';
            document.getElementById('archiveGithubRepoBtn').textContent =
                repo.archived ? 'Restaurar repositório' : 'Arquivar repositório';
            openModal('githubRepoEditModal');
        }

        function openGithubRepoClone(repo) {
            state.cloningGithubRepo = repo;
            document.getElementById('cloneGithubRepoFullName').textContent = repo.full_name;
            document.getElementById('cloneGithubRepoFolderName').value = repo.name || '';
            openModal('githubRepoCloneModal');
        }

        async function githubRepoAction(repo, action, extra = {}) {
            const response = await fetch('/api/github/repos/action', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    repo,
                    action,
                    parent_path: extra.parentPath || '',
                    folder_name: extra.folderName || '',
                    new_name: extra.newName || '',
                    confirmation: extra.confirmation || ''
                })
            });
            const data = await response.json();
            if (!response.ok) throw new Error(data.detail || 'Ação no repositório falhou.');
            appendLog(`GitHub ${action}: ${repo}\n${data.message || ''}`);
            showToast(data.message || 'Operação concluída.');
            return data;
        }

        async function handleGithubRepoCardAction(action, fullName) {
            const repo = findGithubRepo(fullName);
            if (!repo) return;

            if (action === 'copy') {
                try {
                    await navigator.clipboard.writeText(repo.clone_url || repo.url);
                    showToast('URL copiada.');
                } catch (_) {
                    showToast('Não foi possível copiar a URL.', true);
                }
                return;
            }

            if (action === 'edit') {
                openGithubRepoEdit(repo);
                return;
            }

            if (action === 'clone') {
                openGithubRepoClone(repo);
                return;
            }

            if (action === 'select_local') {
                try {
                    await githubRepoAction(repo.full_name, 'select_local');
                    await loadEnvironment({ resetReleaseFields: true });
                    await loadGithubRepositories();
                } catch (error) {
                    showToast(error.message || String(error), true);
                }
                return;
            }

            if (action === 'archive' || action === 'unarchive') {
                const verb = action === 'archive' ? 'arquivar' : 'restaurar';
                if (!window.confirm(`Deseja ${verb} ${repo.full_name}?`)) return;
                try {
                    await githubRepoAction(repo.full_name, action);
                    await loadGithubRepositories();
                } catch (error) {
                    showToast(error.message || String(error), true);
                }
            }
        }

        async function createGithubRepo() {
            const name = document.getElementById('newGithubRepoName').value.trim();
            const visibility = document.getElementById('newGithubRepoVisibility').value;

            if (!name) {
                showToast('Informe o nome do repositório.', true);
                return;
            }

            if (
                visibility === 'public'
                && !window.confirm(`Criar "${name}" como repositório PÚBLICO? O código ficará visível na internet.`)
            ) return;

            const button = document.getElementById('confirmCreateGithubRepoBtn');
            button.disabled = true;
            button.textContent = 'Criando...';

            try {
                const response = await fetch('/api/github/repos/create', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        name,
                        description: document.getElementById('newGithubRepoDescription').value.trim(),
                        visibility,
                        add_readme: document.getElementById('newGithubRepoReadme').checked,
                        clone_after: document.getElementById('newGithubRepoCloneAfter').checked,
                        parent_path: document.getElementById('newGithubRepoParentPath').value.trim()
                    })
                });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Falha ao criar repositório.');

                closeModal('githubRepoCreateModal');
                appendLog(`GitHub create: ${data.repo}\n${data.message || ''}`);
                showToast(data.message || 'Repositório criado.');
                await loadEnvironment({ resetReleaseFields: true });
                await loadGithubRepositories();
            } catch (error) {
                showToast(error.message || String(error), true);
            } finally {
                button.disabled = false;
                button.textContent = 'Criar repositório';
            }
        }

        async function saveGithubRepoEdit() {
            const repo = state.editingGithubRepo;
            if (!repo) return;

            const visibility = document.getElementById('editGithubRepoVisibility').value;
            let confirmPublic = false;

            if (visibility === 'public' && repo.visibility !== 'public') {
                confirmPublic = window.confirm(
                    `ATENÇÃO: ${repo.full_name} ficará PÚBLICO e o código será visível na internet. Confirmar?`
                );
                if (!confirmPublic) return;
            }

            const button = document.getElementById('confirmEditGithubRepoBtn');
            button.disabled = true;
            button.textContent = 'Salvando...';

            try {
                const response = await fetch('/api/github/repos/edit', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        repo: repo.full_name,
                        description: document.getElementById('editGithubRepoDescription').value.trim(),
                        homepage: document.getElementById('editGithubRepoHomepage').value.trim(),
                        visibility,
                        has_issues: document.getElementById('editGithubRepoIssues').checked,
                        has_projects: document.getElementById('editGithubRepoProjects').checked,
                        has_wiki: document.getElementById('editGithubRepoWiki').checked,
                        default_branch: document.getElementById('editGithubRepoDefaultBranch').value.trim(),
                        confirm_public: confirmPublic
                    })
                });
                const data = await response.json();
                if (!response.ok) throw new Error(data.detail || 'Falha ao editar repositório.');

                closeModal('githubRepoEditModal');
                appendLog(`GitHub edit: ${repo.full_name}\n${data.message || ''}`);
                showToast(data.message || 'Repositório atualizado.');
                await loadGithubRepositories();
            } catch (error) {
                showToast(error.message || String(error), true);
            } finally {
                button.disabled = false;
                button.textContent = 'Salvar alterações';
            }
        }

        async function cloneGithubRepo() {
            const repo = state.cloningGithubRepo;
            if (!repo) return;

            const parentPath = document.getElementById('cloneGithubRepoParentPath').value.trim();
            if (!parentPath) {
                showToast('Selecione a pasta-pai para o clone.', true);
                return;
            }

            const button = document.getElementById('confirmCloneGithubRepoBtn');
            button.disabled = true;
            button.textContent = 'Clonando...';

            try {
                await githubRepoAction(repo.full_name, 'clone', {
                    parentPath,
                    folderName: document.getElementById('cloneGithubRepoFolderName').value.trim()
                });
                closeModal('githubRepoCloneModal');
                await loadEnvironment({ resetReleaseFields: true });
                await loadGithubRepositories();
            } catch (error) {
                showToast(error.message || String(error), true);
            } finally {
                button.disabled = false;
                button.textContent = 'Clonar e selecionar';
            }
        }

        async function renameGithubRepo() {
            const repo = state.editingGithubRepo;
            if (!repo) return;

            const newName = document.getElementById('renameGithubRepoName').value.trim();
            if (!newName || newName === repo.name) {
                showToast('Informe um novo nome diferente do atual.', true);
                return;
            }

            const confirmation = window.prompt(
                `Renomear ${repo.full_name} para "${newName}"?\n\nDigite exatamente ${repo.full_name} para confirmar:`
            );
            if (confirmation !== repo.full_name) {
                showToast('Confirmação não corresponde ao repositório.', true);
                return;
            }

            try {
                const data = await githubRepoAction(repo.full_name, 'rename', {
                    newName,
                    confirmation
                });
                closeModal('githubRepoEditModal');
                await loadEnvironment({ resetReleaseFields: true });
                await loadGithubRepositories();
                if (data.local_remote_updated) {
                    appendLog('Origin do clone local atualizado automaticamente após o rename.');
                }
            } catch (error) {
                showToast(error.message || String(error), true);
            }
        }

        async function archiveGithubRepoFromEdit() {
            const repo = state.editingGithubRepo;
            if (!repo) return;
            const action = repo.archived ? 'unarchive' : 'archive';
            const verb = repo.archived ? 'restaurar' : 'arquivar';
            if (!window.confirm(`Deseja ${verb} ${repo.full_name}?`)) return;

            try {
                await githubRepoAction(repo.full_name, action);
                closeModal('githubRepoEditModal');
                await loadGithubRepositories();
            } catch (error) {
                showToast(error.message || String(error), true);
            }
        }

        async function deleteGithubRepo() {
            const repo = state.editingGithubRepo;
            if (!repo) return;

            const confirmation = window.prompt(
                `EXCLUSÃO PERMANENTE.\n\nO repositório ${repo.full_name} será apagado do GitHub.\nDigite exatamente ${repo.full_name} para confirmar:`
            );
            if (confirmation !== repo.full_name) {
                showToast('Exclusão cancelada: confirmação incorreta.', true);
                return;
            }

            if (!window.confirm('Última confirmação: excluir permanentemente este repositório?')) return;

            try {
                await githubRepoAction(repo.full_name, 'delete', { confirmation });
                closeModal('githubRepoEditModal');
                await loadGithubRepositories();
            } catch (error) {
                showToast(error.message || String(error), true);
            }
        }

        async function runIndividualGitCommand(action, options = {}) {
            const payload = {
                action,
                value: options.value || '',
                message: options.message || '',
                create: Boolean(options.create),
                force: Boolean(options.force),
                parent_path: options.parentPath || '',
                folder_name: options.folderName || ''
            };

            gitIndividualResult.classList.add('hidden');

            try {
                const response = await fetch('/api/git-dashboard/command', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });

                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || `Falha em git ${action}.`);
                }

                gitIndividualResult.className =
                    'mt-4 rounded-xl border border-emerald-900/35 bg-emerald-950/15 px-4 py-3 text-xs text-emerald-300';

                gitIndividualResult.innerHTML = `
                    <div class="font-bold">${escapeHtml(data.message || 'Comando concluído.')}</div>
                    <div class="mt-1 break-all font-mono text-[10px] text-slate-500">${escapeHtml(data.command || '')}</div>
                `;
                gitIndividualResult.classList.remove('hidden');

                appendLog(
                    `${data.command || `git ${action}`}` +
                    (data.output ? `\n${data.output}` : '') +
                    (data.message ? `\n${data.message}` : '')
                );

                if (data.selected_project_changed) {
                    showToast('Projeto ativo atualizado.');
                } else {
                    showToast(data.message || 'Comando Git concluído.');
                }

                await loadEnvironment({
                    resetReleaseFields: Boolean(data.selected_project_changed)
                });

                return data;

            } catch (error) {
                gitIndividualResult.className =
                    'mt-4 rounded-xl border border-red-900/40 bg-red-950/20 px-4 py-3 text-xs text-red-300';
                gitIndividualResult.textContent = error.message || String(error);
                gitIndividualResult.classList.remove('hidden');

                appendLog(`ERRO git ${action}: ${error.message || error}`);
                showToast(error.message || `Erro em git ${action}.`, true);
                throw error;
            }
        }

        async function browseCloneParent() {
            const button = document.getElementById('browseCloneParentBtn');
            button.disabled = true;

            try {
                const response = await fetch('/api/git-dashboard/browse-folder', {
                    method: 'POST'
                });
                const data = await response.json();

                if (!response.ok) {
                    throw new Error(data.detail || 'Falha ao selecionar pasta.');
                }

                if (data.path) {
                    document.getElementById('gitCloneParentInput').value = data.path;
                }
            } catch (error) {
                showToast(error.message || 'Erro ao selecionar pasta.', true);
            } finally {
                button.disabled = false;
            }
        }

        document.querySelectorAll('[data-simple-git-action]').forEach(button => {
            button.addEventListener('click', async () => {
                const action = button.dataset.simpleGitAction;
                const originalText = button.textContent;

                button.disabled = true;
                button.textContent = '...';

                try {
                    await runIndividualGitCommand(action);
                } catch (_) {
                    // A mensagem já foi exibida no painel.
                } finally {
                    button.disabled = false;
                    button.textContent = originalText;
                }
            });
        });

        document.getElementById('gitInitBtn').addEventListener('click', async () => {
            try {
                await runIndividualGitCommand('init');
            } catch (_) {}
        });

        document.getElementById('browseCloneParentBtn').addEventListener('click', browseCloneParent);

        document.getElementById('gitCloneBtn').addEventListener('click', async () => {
            const url = document.getElementById('gitCloneUrlInput').value.trim();
            const parentPath = document.getElementById('gitCloneParentInput').value.trim();
            const folderName = document.getElementById('gitCloneFolderInput').value.trim();

            if (!url || !parentPath) {
                showToast('Informe a URL e a pasta-pai do clone.', true);
                return;
            }

            try {
                await runIndividualGitCommand('clone', {
                    value: url,
                    parentPath,
                    folderName
                });
            } catch (_) {}
        });

        document.getElementById('gitAddBtn').addEventListener('click', async () => {
            try {
                await runIndividualGitCommand('add', {
                    value: document.getElementById('gitAddPathInput').value.trim()
                });
            } catch (_) {}
        });

        document.getElementById('gitCommitBtn').addEventListener('click', async () => {
            const message = document.getElementById('gitCommitMessageInput').value.trim();

            if (!message) {
                showToast('Informe a mensagem do commit.', true);
                return;
            }

            try {
                await runIndividualGitCommand('commit', { message });
            } catch (_) {}
        });

        document.getElementById('gitSwitchBtn').addEventListener('click', async () => {
            const branch = document.getElementById('gitBranchNameInput').value.trim();
            if (!branch) {
                showToast('Informe a branch.', true);
                return;
            }
            try {
                await runIndividualGitCommand('switch', { value: branch });
            } catch (_) {}
        });

        document.getElementById('gitSwitchCreateBtn').addEventListener('click', async () => {
            const branch = document.getElementById('gitBranchNameInput').value.trim();
            if (!branch) {
                showToast('Informe a nova branch.', true);
                return;
            }
            try {
                await runIndividualGitCommand('switch', {
                    value: branch,
                    create: true
                });
            } catch (_) {}
        });

        document.getElementById('gitBranchCreateBtn').addEventListener('click', async () => {
            const branch = document.getElementById('gitBranchNameInput').value.trim();
            if (!branch) {
                showToast('Informe a nova branch.', true);
                return;
            }
            try {
                await runIndividualGitCommand('branch_create', { value: branch });
            } catch (_) {}
        });

        document.getElementById('gitBranchDeleteBtn').addEventListener('click', async () => {
            const branch = document.getElementById('gitBranchNameInput').value.trim();
            if (!branch) {
                showToast('Informe a branch que será excluída.', true);
                return;
            }

            const force = document.getElementById('gitBranchForceDeleteInput').checked;
            const confirmation = window.confirm(
                `Excluir a branch local "${branch}"${force ? ' usando -D' : ''}?`
            );

            if (!confirmation) return;

            try {
                await runIndividualGitCommand('branch_delete', {
                    value: branch,
                    force
                });
            } catch (_) {}
        });

        document.getElementById('gitRmCachedBtn').addEventListener('click', async () => {
            const path = document.getElementById('gitRmCachedPathInput').value.trim();

            if (!path) {
                showToast('Informe o arquivo ou pasta.', true);
                return;
            }

            const confirmation = window.confirm(
                `Remover "${path}" somente do índice Git? O arquivo continuará no computador.`
            );
            if (!confirmation) return;

            try {
                await runIndividualGitCommand('rm_cached', { value: path });
            } catch (_) {}
        });


        document.getElementById('refreshGithubReposBtn').addEventListener('click', loadGithubRepositories);
        document.getElementById('createGithubRepoBtn').addEventListener('click', () => {
            document.getElementById('newGithubRepoName').value = '';
            document.getElementById('newGithubRepoDescription').value = '';
            document.getElementById('newGithubRepoVisibility').value = 'private';
            document.getElementById('newGithubRepoReadme').checked = true;
            document.getElementById('newGithubRepoCloneAfter').checked = false;
            document.getElementById('newGithubRepoCloneArea').classList.add('hidden');
            document.getElementById('newGithubRepoParentPath').value = '';
            openModal('githubRepoCreateModal');
        });

        githubRepoSearchInput.addEventListener('input', renderGithubRepos);
        githubRepoVisibilityFilter.addEventListener('change', renderGithubRepos);
        githubRepoStateFilter.addEventListener('change', renderGithubRepos);
        githubRepoSortSelect.addEventListener('change', renderGithubRepos);
        githubRepoScopeSelect.addEventListener('change', loadGithubRepositories);

        document.getElementById('newGithubRepoCloneAfter').addEventListener('change', event => {
            document.getElementById('newGithubRepoCloneArea').classList.toggle('hidden', !event.target.checked);
        });

        document.getElementById('newGithubRepoBrowseParentBtn').addEventListener(
            'click',
            () => browseFolderForInput('newGithubRepoParentPath')
        );
        document.getElementById('cloneGithubRepoBrowseParentBtn').addEventListener(
            'click',
            () => browseFolderForInput('cloneGithubRepoParentPath')
        );

        document.getElementById('confirmCreateGithubRepoBtn').addEventListener('click', createGithubRepo);
        document.getElementById('confirmEditGithubRepoBtn').addEventListener('click', saveGithubRepoEdit);
        document.getElementById('confirmCloneGithubRepoBtn').addEventListener('click', cloneGithubRepo);
        document.getElementById('renameGithubRepoBtn').addEventListener('click', renameGithubRepo);
        document.getElementById('archiveGithubRepoBtn').addEventListener('click', archiveGithubRepoFromEdit);
        document.getElementById('deleteGithubRepoBtn').addEventListener('click', deleteGithubRepo);

        document.querySelectorAll('[data-close-modal]').forEach(button => {
            button.addEventListener('click', () => closeModal(button.dataset.closeModal));
        });

        ['githubRepoCreateModal', 'githubRepoEditModal', 'githubRepoCloneModal'].forEach(id => {
            const modal = document.getElementById(id);
            modal.addEventListener('click', event => {
                if (event.target === modal) closeModal(id);
            });
        });

        gitCreateExecutableBtn.addEventListener(
            'click',
            openGitBuildExecutableModal
        );
        gitShowReadmeBtn.addEventListener(
            'click',
            openGitReadmeModal
        );

        document.getElementById('gitCloseBuildExecutableModalBtn').addEventListener(
            'click',
            closeGitBuildExecutableModal
        );
        document.getElementById('gitCancelBuildModalBtn').addEventListener(
            'click',
            closeGitBuildExecutableModal
        );
        gitRunExecutableBuildBtn.addEventListener(
            'click',
            runGitExecutableBuild
        );
        gitOpenBuildFolderBtn.addEventListener(
            'click',
            openGitBuildOutputFolder
        );

        gitBuildExecutableModal.addEventListener('click', event => {
            if (event.target === gitBuildExecutableModal) {
                closeGitBuildExecutableModal();
            }
        });

        document.getElementById('gitCloseReadmeModalBtn').addEventListener(
            'click',
            closeGitReadmeModal
        );
        gitReadmeModal.addEventListener('click', event => {
            if (event.target === gitReadmeModal) {
                closeGitReadmeModal();
            }
        });

        document.addEventListener('keydown', event => {
            if (event.key !== 'Escape') return;
            closeGitBuildExecutableModal();
            closeGitReadmeModal();
        });

        document.getElementById('refreshGitStatusBtn').addEventListener('click', () => loadEnvironment());
        document.getElementById('chooseProjectFolderBtn').addEventListener('click', chooseProjectFolder);
        document.getElementById('forgetProjectBtn').addEventListener('click', forgetActiveProject);
        projectSelect.addEventListener('change', () => activateProject(projectSelect.value));
        document.getElementById('runSecurityAuditBtn').addEventListener('click', runSecurityAudit);
        document.getElementById('publishGitBtn').addEventListener('click', publishProject);
        document.getElementById('createReleaseBtn').addEventListener('click', createRelease);

        document.getElementById('clearGitLogBtn').addEventListener('click', () => {
            gitActionLog.textContent = 'Pronto.';
        });

        const gitSidebar = document.getElementById('gitSidebar');
        const gitSidebarOverlay = document.getElementById('gitSidebarOverlay');

        document.getElementById('gitMenuBtn').addEventListener('click', () => {
            gitSidebar.classList.add('open');
            gitSidebarOverlay.classList.remove('hidden');
        });

        gitSidebarOverlay.addEventListener('click', () => {
            gitSidebar.classList.remove('open');
            gitSidebarOverlay.classList.add('hidden');
        });

        loadEnvironment({ resetReleaseFields: true });
        loadGithubRepositories();
    </script>
</body>
</html>
"""


# ============================================================
# ROTAS
# ============================================================

@app.get("/", response_class=HTMLResponse, include_in_schema=False)
def home() -> HTMLResponse:
    """Entrega a interface web unificada."""
    return HTMLResponse(content=INDEX_HTML)


@app.get("/apps", response_class=HTMLResponse, include_in_schema=False)
def apps_gallery() -> HTMLResponse:
    """Entrega a galeria simplificada de aplicativos (ícone + nome)."""
    return HTMLResponse(content=APPS_HTML)


@app.get("/links", response_class=HTMLResponse, include_in_schema=False)
def favorite_links_page() -> HTMLResponse:
    """Entrega a biblioteca de links favoritos organizada por categorias."""
    return HTMLResponse(content=LINKS_HTML)


@app.get("/local-apps", response_class=HTMLResponse, include_in_schema=False)
def local_apps_page() -> HTMLResponse:
    """Entrega a central de aplicativos instalados no computador."""
    return HTMLResponse(content=LOCAL_APPS_HTML)


@app.get("/windows-diagnostics", response_class=HTMLResponse, include_in_schema=False)
def windows_diagnostics_page() -> HTMLResponse:
    """Entrega a central de diagnóstico e manutenção do Windows."""
    return HTMLResponse(content=WINDOWS_DIAGNOSTICS_HTML)


@app.get("/workspace", response_class=HTMLResponse, include_in_schema=False)
def workspace_page() -> HTMLResponse:
    """Entrega a área personalizada de projetos."""
    return HTMLResponse(content=WORKSPACE_HTML)


@app.get("/git-github", response_class=HTMLResponse, include_in_schema=False)
def git_github_page() -> HTMLResponse:
    """Entrega o dashboard local de Git e GitHub."""
    return HTMLResponse(content=GIT_GITHUB_HTML)


@app.get("/api/workspace/explorer")
def workspace_explorer_api(request: Request) -> dict[str, Any]:
    """Lista projetos reais da raiz local do Workspace."""
    _local_request_only(request)
    return workspace_explorer_snapshot()


@app.post("/api/workspace/explorer/action")
def workspace_explorer_action_api(
    payload: WorkspaceFolderActionRequest,
    request: Request,
) -> dict[str, Any]:
    """Executa ações seguras sobre uma pasta listada no Workspace."""
    _local_request_only(request)
    return workspace_folder_action(payload)


@app.get("/api/projects")
def list_projects() -> list[dict[str, Any]]:
    """Lista projetos e recursos associados."""
    with PROJECTS_LOCK:
        return read_projects()


@app.post("/api/projects", status_code=status.HTTP_201_CREATED)
def create_project(payload: ProjectCreate) -> dict[str, Any]:
    """Cria um projeto personalizado."""
    project_data = normalized_project_payload(payload)

    with PROJECTS_LOCK:
        projects = read_projects()
        next_id = max(
            (int(project.get("id", 0)) for project in projects),
            default=0,
        ) + 1

        saved = {
            "id": next_id,
            **project_data,
            "resources": [],
            "phases": default_project_phases(),
        }

        projects.append(saved)
        write_projects(projects)

    return saved


@app.put("/api/projects/{project_id}")
def update_project(project_id: int, payload: ProjectCreate) -> dict[str, Any]:
    """Edita dados do projeto preservando suas ferramentas."""
    project_data = normalized_project_payload(payload)

    with PROJECTS_LOCK:
        projects = read_projects()
        index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        resources = projects[index].get("resources", [])
        phases = projects[index].get("phases", default_project_phases())
        saved = {
            "id": project_id,
            **project_data,
            "resources": resources,
            "phases": phases,
        }

        projects[index] = saved
        write_projects(projects)

    return saved


@app.delete("/api/projects/{project_id}")
def delete_project(project_id: int) -> dict[str, Any]:
    """Exclui um projeto."""
    with PROJECTS_LOCK:
        projects = read_projects()
        project = next(
            (item for item in projects if int(item.get("id", -1)) == project_id),
            None,
        )

        if project is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        write_projects([
            item for item in projects
            if int(item.get("id", -1)) != project_id
        ])

    return {
        "message": "Projeto excluído.",
        "deleted_id": project_id,
    }


@app.post("/api/projects/{project_id}/phases", status_code=status.HTTP_201_CREATED)
def add_project_phase(
    project_id: int,
    payload: ProjectPhaseCreate,
) -> dict[str, Any]:
    """Adiciona uma fase ao projeto."""
    phase_name = payload.name.strip()
    if not phase_name:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O nome da fase é obrigatório.",
        )

    progress = 100 if payload.status == "concluido" else int(payload.progress)

    with PROJECTS_LOCK:
        projects = read_projects()
        index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        phases = projects[index].setdefault("phases", [])
        next_id = max(
            (int(phase.get("id", 0)) for phase in phases if isinstance(phase, dict)),
            default=0,
        ) + 1

        phases.append({
            "id": next_id,
            "name": phase_name,
            "status": payload.status,
            "progress": progress,
        })

        write_projects(projects)
        return projects[index]


@app.put("/api/projects/{project_id}/phases/{phase_id}")
def update_project_phase(
    project_id: int,
    phase_id: int,
    payload: ProjectPhaseCreate,
) -> dict[str, Any]:
    """Edita uma fase e seu percentual de andamento."""
    phase_name = payload.name.strip()
    if not phase_name:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O nome da fase é obrigatório.",
        )

    progress = 100 if payload.status == "concluido" else int(payload.progress)

    with PROJECTS_LOCK:
        projects = read_projects()
        project_index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if project_index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        phases = projects[project_index].setdefault("phases", [])
        phase_index = next(
            (i for i, phase in enumerate(phases) if int(phase.get("id", -1)) == phase_id),
            None,
        )

        if phase_index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Fase não encontrada.",
            )

        phases[phase_index] = {
            "id": phase_id,
            "name": phase_name,
            "status": payload.status,
            "progress": progress,
        }

        write_projects(projects)
        return projects[project_index]


@app.delete("/api/projects/{project_id}/phases/{phase_id}")
def delete_project_phase(
    project_id: int,
    phase_id: int,
) -> dict[str, Any]:
    """Exclui uma fase do projeto."""
    with PROJECTS_LOCK:
        projects = read_projects()
        project_index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if project_index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        phases = projects[project_index].setdefault("phases", [])
        original_length = len(phases)
        projects[project_index]["phases"] = [
            phase for phase in phases
            if int(phase.get("id", -1)) != phase_id
        ]

        if len(projects[project_index]["phases"]) == original_length:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Fase não encontrada.",
            )

        write_projects(projects)
        return projects[project_index]


@app.post("/api/projects/{project_id}/resources")
def add_project_resource(
    project_id: int,
    payload: ProjectResourceRequest,
) -> dict[str, Any]:
    """Associa App, App Local ou Link ao projeto."""
    normalized_id = normalize_project_resource_id(
        payload.resource_type,
        payload.resource_id,
    )

    if not project_resource_exists(payload.resource_type, normalized_id):
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="O recurso selecionado não existe mais.",
        )

    with PROJECTS_LOCK:
        projects = read_projects()
        index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        resources = projects[index].setdefault("resources", [])
        key = project_resource_key(payload.resource_type, normalized_id)

        exists = any(
            project_resource_key(
                str(item.get("resource_type", "")),
                item.get("resource_id", ""),
            ) == key
            for item in resources
            if isinstance(item, dict)
            and str(item.get("resource_type", "")) in {"app", "local_app", "link"}
        )

        if not exists:
            resources.append({
                "resource_type": payload.resource_type,
                "resource_id": normalized_id,
            })
            write_projects(projects)

        return projects[index]


@app.post("/api/projects/{project_id}/resources/import/{source_project_id}")
def import_project_resources(
    project_id: int,
    source_project_id: int,
) -> dict[str, Any]:
    """Copia recursos de outro projeto, sem duplicar os já vinculados."""
    if project_id == source_project_id:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O projeto de origem deve ser diferente do projeto atual.",
        )

    with PROJECTS_LOCK:
        projects = read_projects()

        target = next(
            (project for project in projects if int(project.get("id", -1)) == project_id),
            None,
        )
        source = next(
            (project for project in projects if int(project.get("id", -1)) == source_project_id),
            None,
        )

        if target is None or source is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto de origem ou destino não encontrado.",
            )

        source_resources = [
            dict(resource)
            for resource in source.get("resources", [])
            if isinstance(resource, dict)
        ]

    valid_resources: list[dict[str, Any]] = []
    missing = 0

    for resource in source_resources:
        resource_type = str(resource.get("resource_type", ""))
        if resource_type not in {"app", "local_app", "link"}:
            missing += 1
            continue

        try:
            normalized_id = normalize_project_resource_id(
                resource_type,
                resource.get("resource_id", ""),
            )
        except HTTPException:
            missing += 1
            continue

        if not project_resource_exists(resource_type, normalized_id):
            missing += 1
            continue

        valid_resources.append({
            "resource_type": resource_type,
            "resource_id": normalized_id,
        })

    with PROJECTS_LOCK:
        projects = read_projects()
        target_index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if target_index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto de destino não encontrado.",
            )

        target_resources = projects[target_index].setdefault("resources", [])

        existing_keys = {
            project_resource_key(
                str(resource.get("resource_type", "")),
                resource.get("resource_id", ""),
            )
            for resource in target_resources
            if isinstance(resource, dict)
            and str(resource.get("resource_type", "")) in {"app", "local_app", "link"}
        }

        imported = 0
        duplicates = 0

        for resource in valid_resources:
            key = project_resource_key(
                resource["resource_type"],
                resource["resource_id"],
            )

            if key in existing_keys:
                duplicates += 1
                continue

            target_resources.append(resource)
            existing_keys.add(key)
            imported += 1

        if imported:
            write_projects(projects)

        target_project = projects[target_index]

    return {
        "project": target_project,
        "imported": imported,
        "duplicates": duplicates,
        "missing": missing,
    }


@app.delete("/api/projects/{project_id}/resources/{resource_type}/{resource_id}")
def remove_project_resource(
    project_id: int,
    resource_type: str,
    resource_id: str,
) -> dict[str, Any]:
    """Remove App, App Local ou Link de um projeto."""
    normalized_id = normalize_project_resource_id(resource_type, resource_id)
    target_key = project_resource_key(resource_type, normalized_id)

    with PROJECTS_LOCK:
        projects = read_projects()
        index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        resources = projects[index].setdefault("resources", [])
        kept: list[dict[str, Any]] = []

        for item in resources:
            if not isinstance(item, dict):
                continue

            item_type = str(item.get("resource_type", ""))
            try:
                item_key = project_resource_key(
                    item_type,
                    item.get("resource_id", ""),
                )
            except HTTPException:
                kept.append(item)
                continue

            if item_key != target_key:
                kept.append(item)

        projects[index]["resources"] = kept
        write_projects(projects)
        return projects[index]


@app.put("/api/projects/{project_id}/resources/order")
def reorder_project_resources(
    project_id: int,
    payload: ProjectResourcesOrderRequest,
) -> dict[str, Any]:
    """Reserva endpoint para ordenação persistente dos recursos do projeto."""
    with PROJECTS_LOCK:
        projects = read_projects()
        index = next(
            (i for i, project in enumerate(projects) if int(project.get("id", -1)) == project_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Projeto não encontrado.",
            )

        current_items = [
            item for item in projects[index].get("resources", [])
            if isinstance(item, dict)
            and str(item.get("resource_type", "")) in {"app", "local_app", "link"}
        ]

        current = {
            project_resource_key(
                str(item.get("resource_type")),
                item.get("resource_id", ""),
            )
            for item in current_items
        }

        requested_normalized = [
            (
                item.resource_type,
                normalize_project_resource_id(
                    item.resource_type,
                    item.resource_id,
                ),
            )
            for item in payload.resources
        ]

        requested_keys = {
            project_resource_key(resource_type, resource_id)
            for resource_type, resource_id in requested_normalized
        }

        if requested_keys != current or len(requested_normalized) != len(current_items):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="A lista de recursos mudou. Atualize a página e tente novamente.",
            )

        projects[index]["resources"] = [
            {"resource_type": resource_type, "resource_id": resource_id}
            for resource_type, resource_id in requested_normalized
        ]

        write_projects(projects)
        return projects[index]


@app.get("/api/windows-diagnostics/status")
def windows_diagnostics_status() -> dict[str, Any]:
    """Retorna disponibilidade das ferramentas e navegadores."""
    return get_windows_diagnostics_status()


@app.post("/api/windows-diagnostics/disk-cleanup/scan")
def diagnostics_disk_cleanup_scan(
    payload: DiskCleanupRequest,
    request: Request,
) -> dict[str, Any]:
    """Analisa somente as áreas de limpeza conhecidas pelo Hub."""
    require_local_windows(request)

    valid_areas = [
        area for area in payload.areas
        if area in DISK_CLEANUP_AREA_LABELS
    ]

    if not valid_areas:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nenhuma área válida foi selecionada.",
        )

    return scan_disk_cleanup_areas(valid_areas)


@app.post("/api/windows-diagnostics/disk-cleanup/clean")
def diagnostics_disk_cleanup_clean(
    payload: DiskCleanupRequest,
    request: Request,
) -> dict[str, Any]:
    """Limpa as áreas selecionadas depois de confirmação explícita."""
    require_local_windows(request)

    if not payload.confirmed:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="A limpeza exige confirmação explícita.",
        )

    valid_areas = [
        area for area in payload.areas
        if area in DISK_CLEANUP_AREA_LABELS
    ]

    if not valid_areas:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nenhuma área válida foi selecionada.",
        )

    return clean_disk_cleanup_areas(valid_areas)


@app.post("/api/windows-diagnostics/temp-locations")
def diagnostics_open_temp_locations(request: Request) -> dict[str, Any]:
    """Abre Prefetch, TMP e TEMP no Explorador."""
    require_local_windows(request)

    try:
        opened = open_windows_temp_locations()
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir os diretórios: {exc}",
        ) from exc

    if not opened:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Nenhum dos diretórios temporários foi encontrado.",
        )

    return {
        "message": f"{len(opened)} local(is) aberto(s) no Explorador.",
        "opened": opened,
    }


@app.post("/api/windows-diagnostics/performance-report")
def diagnostics_performance_report(request: Request) -> dict[str, str]:
    """Inicia perfmon /report."""
    require_local_windows(request)

    try:
        launch_windows_tool(
            "perfmon.exe",
            "/report",
            elevated=True,
        )
    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="perfmon.exe não foi encontrado.",
        )
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível iniciar o relatório: {exc}",
        ) from exc

    return {
        "message": "Relatório solicitado com privilégios administrativos. Confirme o UAC do Windows e aguarde a coleta.",
    }


@app.post("/api/windows-diagnostics/memory-diagnostic")
def diagnostics_memory_diagnostic(request: Request) -> dict[str, str]:
    """Abre mdsched.exe; o usuário decide se/quando reiniciar."""
    require_local_windows(request)

    try:
        launch_windows_tool(
            "mdsched.exe",
            elevated=True,
        )
    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="mdsched.exe não foi encontrado.",
        )
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir o diagnóstico de memória: {exc}",
        ) from exc

    return {
        "message": "Diagnóstico de Memória solicitado. Confirme o UAC do Windows; o reinício só ocorrerá se você escolher essa opção.",
    }


@app.post("/api/windows-diagnostics/browser-cleanup/{browser_id}")
def diagnostics_browser_cleanup(
    browser_id: str,
    request: Request,
) -> dict[str, str]:
    """Abre a interface nativa de limpeza do navegador escolhido."""
    require_local_windows(request)

    try:
        browser_name = open_browser_cleanup(browser_id)
    except KeyError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Navegador não reconhecido.",
        )
    except FileNotFoundError:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="O navegador selecionado não foi encontrado neste computador.",
        )
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir o navegador: {exc}",
        ) from exc

    return {
        "message": f"Tela de limpeza do {browser_name} aberta.",
    }


@app.get("/api/local-apps")
def list_local_apps() -> list[dict[str, Any]]:
    """Detecta aplicativos locais conhecidos e informa disponibilidade."""
    return get_local_apps_status()


@app.post("/api/local-apps/browse-exe")
def browse_local_executable(request: Request) -> dict[str, str]:
    """Abre o seletor nativo do Windows para escolher um executável."""
    client_host = request.client.host if request.client else ""

    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="O seletor de executável só pode ser usado localmente.",
        )

    if os.name != "nt":
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="O seletor de executável está disponível somente no Windows.",
        )

    script = r"""
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Filter = "Aplicativos (*.exe)|*.exe"
$dialog.Title = "Selecionar aplicativo"
$dialog.Multiselect = $false
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.FileName
}
"""

    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-STA", "-Command", script],
            capture_output=True,
            text=True,
            timeout=120,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError) as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir o seletor: {exc}",
        ) from exc

    path = result.stdout.strip()

    if not path:
        return {"path": "", "suggested_name": ""}

    selected = Path(path)

    return {
        "path": str(selected),
        "suggested_name": selected.stem,
    }


@app.post(
    "/api/local-apps/custom",
    status_code=status.HTTP_201_CREATED,
)
def add_custom_local_app(
    new_app: CustomLocalAppCreate,
    request: Request,
) -> dict[str, Any]:
    """Cadastra manualmente um aplicativo local por caminho de executável."""
    client_host = request.client.host if request.client else ""

    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Aplicativos locais só podem ser cadastrados neste computador.",
        )

    executable = Path(
        os.path.expandvars(new_app.executable_path.strip().strip('"'))
    ).expanduser()

    if not executable.is_file():
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O executável informado não foi encontrado.",
        )

    if executable.suffix.casefold() != ".exe":
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Selecione um arquivo executável .exe.",
        )

    clean_name = new_app.name.strip()
    clean_description = new_app.description.strip()
    clean_icon = new_app.icon.strip() or "🖥️"

    with LOCAL_APPS_LOCK:
        items = read_custom_local_apps()

        normalized_path = os.path.normcase(str(executable.resolve()))

        duplicate = any(
            os.path.normcase(
                str(Path(str(item.get("custom_path", ""))).expanduser())
            ) == normalized_path
            for item in items
            if item.get("custom_path")
        )

        if duplicate:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Esse executável já está cadastrado em Apps Locais.",
            )

        saved = {
            "id": f"custom-{uuid4().hex}",
            "name": clean_name,
            "group": new_app.group,
            "icon": clean_icon,
            "description": clean_description or "Aplicativo local adicionado pelo usuário.",
            "custom_path": str(executable.resolve()),
            "custom": True,
        }

        items.append(saved)
        write_custom_local_apps(items)

    return {
        **saved,
        "installed": True,
        "source": "Adicionado por você",
    }


@app.delete("/api/local-apps/custom/{app_id}")
def delete_custom_local_app(app_id: str, request: Request) -> dict[str, Any]:
    """Remove apenas cadastros locais criados pelo usuário."""
    client_host = request.client.host if request.client else ""

    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Aplicativos locais só podem ser removidos neste computador.",
        )

    with LOCAL_APPS_LOCK:
        items = read_custom_local_apps()

        item = next(
            (entry for entry in items if entry.get("id") == app_id),
            None,
        )

        if item is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Aplicativo local personalizado não encontrado.",
            )

        write_custom_local_apps([
            entry for entry in items
            if entry.get("id") != app_id
        ])

    return {
        "message": "Aplicativo local removido.",
        "deleted_id": app_id,
        "deleted_name": item.get("name"),
    }


@app.post("/api/local-apps/{app_id}/launch")
def launch_local_app_endpoint(app_id: str, request: Request) -> dict[str, Any]:
    """Abre um aplicativo local da lista segura do Hub."""
    client_host = request.client.host if request.client else ""

    if client_host not in {"127.0.0.1", "::1", "localhost", "testclient"}:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Aplicativos locais só podem ser abertos a partir deste computador.",
        )

    return launch_catalog_app(app_id)


@app.get("/api/user-profile")
def get_user_profile() -> dict[str, str]:
    """Retorna o perfil local exibido na interface."""
    with USER_PROFILE_LOCK:
        return public_user_profile(read_user_profile())


@app.post("/api/user-profile/image", status_code=status.HTTP_201_CREATED)
async def upload_user_profile_image(request: Request) -> dict[str, str]:
    """Atualiza a imagem de perfil do usuário local."""
    content_type = request.headers.get("content-type", "").split(";", 1)[0].strip().lower()
    extension = ALLOWED_ICON_TYPES.get(content_type)

    if extension is None:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail="Formato não suportado. Use PNG, JPG, WEBP ou GIF.",
        )

    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > MAX_PROFILE_IMAGE_SIZE:
                raise HTTPException(
                    status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                    detail="A imagem do perfil deve ter no máximo 5 MB.",
                )
        except ValueError:
            pass

    data = await request.body()
    if not data:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O arquivo da imagem do perfil está vazio.",
        )
    if len(data) > MAX_PROFILE_IMAGE_SIZE:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail="A imagem do perfil deve ter no máximo 5 MB.",
        )

    USER_MEDIA_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"profile_{uuid4().hex}{extension}"
    destination = USER_MEDIA_DIR / filename

    try:
        destination.write_bytes(data)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível salvar a imagem do perfil: {exc}",
        ) from exc

    with USER_PROFILE_LOCK:
        profile = read_user_profile()
        old_filename = str(profile.get("image_filename", "") or "").strip()

        profile["image"] = f"/user-media/{filename}"
        profile["image_filename"] = filename
        write_user_profile(profile)

    if old_filename:
        old_file = USER_MEDIA_DIR / Path(old_filename).name
        if old_file.is_file() and old_file.name != filename:
            try:
                old_file.unlink()
            except OSError:
                pass

    return public_user_profile(profile)


@app.get("/user-media/{filename}", include_in_schema=False)
def get_user_media(filename: str) -> FileResponse:
    """Entrega arquivos de mídia do perfil do usuário."""
    safe_name = Path(filename).name
    if safe_name != filename:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Arquivo não encontrado.")

    media_file = USER_MEDIA_DIR / safe_name
    if not media_file.is_file():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Arquivo não encontrado.")

    return FileResponse(media_file)


@app.get("/brand-image", include_in_schema=False)
def get_brand_image() -> FileResponse:
    """Entrega a arte da marca usada no layout."""
    ensure_brand_image_file()

    if not BRAND_IMAGE_FILE.is_file():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Imagem da marca não encontrada.")

    return FileResponse(BRAND_IMAGE_FILE)


@app.get("/favicon.ico", include_in_schema=False)
def get_favicon() -> FileResponse:
    """Usa a própria marca do TECH TOOL HUB como favicon."""
    ensure_brand_image_file()

    if not BRAND_IMAGE_FILE.is_file():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Favicon não encontrado.",
        )

    return FileResponse(BRAND_IMAGE_FILE, media_type="image/png")


@app.post("/api/icons", status_code=status.HTTP_201_CREATED)
async def upload_icon(request: Request) -> dict[str, str]:
    """Recebe o arquivo do ícone como corpo binário e salva localmente."""
    content_type = request.headers.get("content-type", "").split(";", 1)[0].strip().lower()
    extension = ALLOWED_ICON_TYPES.get(content_type)

    if extension is None:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail="Formato não suportado. Use PNG, JPG, WEBP ou GIF.",
        )

    content_length = request.headers.get("content-length")
    if content_length:
        try:
            if int(content_length) > MAX_ICON_SIZE:
                raise HTTPException(
                    status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                    detail="O ícone deve ter no máximo 2 MB.",
                )
        except ValueError:
            pass

    data = await request.body()
    if not data:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O arquivo de ícone está vazio.",
        )
    if len(data) > MAX_ICON_SIZE:
        raise HTTPException(
            status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
            detail="O ícone deve ter no máximo 2 MB.",
        )

    ICONS_DIR.mkdir(parents=True, exist_ok=True)
    filename = f"{uuid4().hex}{extension}"
    destination = ICONS_DIR / filename

    try:
        destination.write_bytes(data)
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível salvar o ícone: {exc}",
        ) from exc

    return {"icon": f"/icons/{filename}"}


@app.get("/icons/{filename}", include_in_schema=False)
def get_icon(filename: str) -> FileResponse:
    """Entrega um ícone previamente enviado para o Hub."""
    safe_name = Path(filename).name
    if safe_name != filename:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Ícone não encontrado.")

    icon_file = ICONS_DIR / safe_name
    if not icon_file.is_file():
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="Ícone não encontrado.")

    return FileResponse(icon_file)


@app.post("/api/git-dashboard/select-folder")
def git_dashboard_select_folder(
    request: Request,
) -> dict[str, Any]:
    """Abre o seletor nativo e ativa a pasta escolhida."""
    _local_request_only(request)
    selected_path = _choose_project_folder()

    if not selected_path:
        return {
            "cancelled": True,
            "selected_path": "",
            "repo_root": "",
            "is_git": False,
        }

    result = _remember_git_project(
        selected_path,
        activate=True,
    )
    return {
        "cancelled": False,
        **result,
    }


@app.post("/api/git-dashboard/activate-project")
def git_dashboard_activate_project(
    payload: GitProjectPathRequest,
    request: Request,
) -> dict[str, Any]:
    """Ativa um projeto já conhecido ou informado manualmente."""
    _local_request_only(request)
    return _remember_git_project(
        payload.path,
        activate=True,
    )


@app.post("/api/git-dashboard/forget-project")
def git_dashboard_forget_project(
    payload: GitProjectPathRequest,
    request: Request,
) -> dict[str, Any]:
    """Remove um projeto da lista sem apagar qualquer arquivo."""
    _local_request_only(request)

    target = _clean_project_path(payload.path)
    state = _load_git_projects_state()

    if target is None:
        return {
            "removed": False,
            "selected_path": state.get("selected_path", ""),
        }

    target_text = str(target)
    projects = [
        item
        for item in state.get("projects", [])
        if (
            str(item.get("path", "")).lower()
            != target_text.lower()
            if os.name == "nt"
            else str(item.get("path", "")) != target_text
        )
    ]

    selected_text = str(state.get("selected_path", ""))
    was_selected = (
        selected_text.lower() == target_text.lower()
        if os.name == "nt"
        else selected_text == target_text
    )

    state["projects"] = projects

    if was_selected:
        state["selected_path"] = (
            str(projects[0].get("path", ""))
            if projects
            else ""
        )

    _save_git_projects_state(state)

    return {
        "removed": True,
        "selected_path": state.get("selected_path", ""),
    }


@app.get("/api/security-policy")
def get_security_policy(request: Request) -> dict[str, Any]:
    """Expõe a política de segurança sem dados do usuário."""
    _local_request_only(request)
    return security_policy_payload()


@app.get("/api/github/repos")
def github_repositories(
    request: Request,
    scope: str = "owner",
) -> dict[str, Any]:
    """Lista repositórios da conta GitHub autenticada."""
    _local_request_only(request)
    return list_github_repositories(scope)


@app.post("/api/github/repos/create")
def github_repository_create(
    payload: GitHubRepoCreateRequest,
    request: Request,
) -> dict[str, Any]:
    """Cria um repositório na conta GitHub autenticada."""
    _local_request_only(request)
    return create_github_repository(payload)


@app.post("/api/github/repos/edit")
def github_repository_edit(
    payload: GitHubRepoEditRequest,
    request: Request,
) -> dict[str, Any]:
    """Edita metadados/configurações de um repositório."""
    _local_request_only(request)
    return edit_github_repository(payload)


@app.post("/api/github/repos/action")
def github_repository_action(
    payload: GitHubRepoActionRequest,
    request: Request,
) -> dict[str, Any]:
    """Clone, seleção local, archive, rename e delete."""
    _local_request_only(request)
    return manage_github_repository(payload)


@app.post("/api/git-dashboard/browse-folder")
def git_dashboard_browse_folder(
    request: Request,
) -> dict[str, str]:
    """Escolhe uma pasta sem alterar o projeto ativo."""
    _local_request_only(request)
    selected_path = _choose_project_folder()
    return {"path": selected_path}


@app.post("/api/git-dashboard/command")
def git_dashboard_command(
    payload: GitCommandRequest,
    request: Request,
) -> dict[str, Any]:
    """Executa um comando Git individual previamente autorizado."""
    _local_request_only(request)
    return execute_individual_git_command(payload)


@app.get("/api/git-dashboard/status")
def git_dashboard_status(request: Request) -> dict[str, Any]:
    """Resumo do projeto ativo, Git, origin, GitHub CLI e Release."""
    _local_request_only(request)
    return git_github_environment()


@app.post("/api/git-dashboard/security-check")
def git_dashboard_security_check(request: Request) -> dict[str, Any]:
    """Audita possíveis vazamentos antes de publicar."""
    _local_request_only(request)
    return run_git_security_audit()


@app.post("/api/git-dashboard/publish")
def git_dashboard_publish(
    payload: GitPublishRequest,
    request: Request,
) -> dict[str, Any]:
    """Executa auditoria, git add, commit e push."""
    _local_request_only(request)
    return publish_project_to_github(payload.commit_message)


@app.post("/api/git-dashboard/release")
def git_dashboard_release(
    payload: GitReleaseRequest,
    request: Request,
) -> dict[str, Any]:
    """Cria GitHub Release usando gh CLI autenticado."""
    _local_request_only(request)
    return create_github_release(payload)


@app.get("/api/readme")
def get_project_readme(request: Request) -> dict[str, str]:
    """Exibe o README do projeto dentro do Dashboard."""
    _local_request_only(request)
    return {"content": read_project_readme()}


@app.get("/api/build/status")
def get_build_status(request: Request) -> dict[str, Any]:
    """Estado do build iniciado pelo Dashboard."""
    _local_request_only(request)
    return public_build_state()


@app.post("/api/build/executable", status_code=status.HTTP_202_ACCEPTED)
def start_executable_build(request: Request) -> dict[str, Any]:
    """Inicia a criação de TechToolHub.exe em background."""
    _local_request_only(request)

    environment = build_environment_status()
    if not environment["available"]:
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=" ".join(environment["reasons"]),
        )

    with BUILD_LOCK:
        if BUILD_STATE.get("running"):
            return public_build_state()

        BUILD_STATE.update({
            "running": True,
            "status": "starting",
            "message": "Preparando o build...",
            "started_at": "",
            "finished_at": "",
            "output_path": "",
            "log": ["Preparando o build do TECH TOOL HUB..."],
        })

    thread = threading.Thread(
        target=_run_executable_build,
        name="TechToolHubBuild",
        daemon=True,
    )
    thread.start()
    return public_build_state()


@app.post("/api/build/open-output")
def open_build_output(request: Request) -> dict[str, str]:
    """Abre a pasta onde o EXE foi criado."""
    _local_request_only(request)

    if os.name != "nt":
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="Esta ação está disponível no Windows.",
        )

    if not BUILD_RELEASE_DIR.is_dir():
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="A pasta de saída ainda não existe.",
        )

    try:
        os.startfile(str(BUILD_RELEASE_DIR))  # type: ignore[attr-defined]
    except OSError as exc:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Não foi possível abrir a pasta: {exc}",
        ) from exc

    return {"message": f"Pasta aberta: {BUILD_RELEASE_DIR}"}


@app.get("/api/links/browser-import/status")
def browser_import_status(request: Request) -> dict[str, Any]:
    """Leitura de favoritos desativada pela política zero-credenciais."""
    _local_request_only(request)
    raise HTTPException(
        status_code=status.HTTP_410_GONE,
        detail=(
            "A leitura de favoritos do navegador foi desativada. "
            "O TECH TOOL HUB não coleta nem persiste links pessoais."
        ),
    )


@app.post("/api/links/browser-import")
def import_browser_bookmarks(
    payload: BrowserBookmarksImportRequest,
    request: Request,
) -> dict[str, Any]:
    """Importação desativada pela política zero-credenciais."""
    _local_request_only(request)
    raise HTTPException(
        status_code=status.HTTP_410_GONE,
        detail=(
            "A importação de favoritos do navegador foi desativada. "
            "Links pessoais não são armazenados pelo TECH TOOL HUB."
        ),
    )


@app.get("/api/links", response_model=list[FavoriteLinkItem])
def list_favorite_links() -> list[dict[str, Any]]:
    """Lista todos os links favoritos na ordem persistida."""
    with LINKS_LOCK:
        return read_links()


@app.post(
    "/api/links",
    response_model=FavoriteLinkItem,
    status_code=status.HTTP_201_CREATED,
)
def add_favorite_link(new_link: FavoriteLinkCreate) -> dict[str, Any]:
    """Adiciona um link favorito e gera o ID automaticamente."""
    validated_url = validate_http_url(new_link.url)
    link_data = model_to_dict(new_link)

    link_data["title"] = link_data["title"].strip()
    link_data["category"] = link_data["category"].strip()
    link_data["url"] = validated_url
    link_data["note"] = link_data.get("note", "").strip()

    if not link_data["title"] or not link_data["category"]:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome e categoria não podem conter apenas espaços.",
        )

    with LINKS_LOCK:
        links = read_links()

        normalized_url = validated_url.rstrip("/").casefold()
        if any(
            str(item.get("url", "")).strip().rstrip("/").casefold() == normalized_url
            for item in links
            if isinstance(item, dict)
        ):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Este link já está cadastrado nos favoritos.",
            )

        next_id = max(
            (int(item.get("id", 0)) for item in links if isinstance(item, dict)),
            default=0,
        ) + 1

        saved = {"id": next_id, **link_data}
        links.append(saved)
        write_links(links)

    return saved


@app.put("/api/links/order")
def reorder_favorite_links(order: LinksOrderRequest) -> dict[str, Any]:
    """Salva a ordem definida por arrastar e soltar."""
    if not order.ids:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A ordem dos links não pode estar vazia.",
        )

    with LINKS_LOCK:
        links = read_links()

        current_ids = [int(item.get("id", -1)) for item in links]
        requested_ids = [int(link_id) for link_id in order.ids]

        if len(requested_ids) != len(set(requested_ids)):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="A nova ordem contém IDs duplicados.",
            )

        if set(requested_ids) != set(current_ids) or len(requested_ids) != len(current_ids):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="A lista de links mudou. Atualize a página e tente novamente.",
            )

        links_by_id = {int(item["id"]): item for item in links}
        reordered = [links_by_id[link_id] for link_id in requested_ids]
        write_links(reordered)

    return {
        "message": "Ordem dos links salva com sucesso.",
        "ids": requested_ids,
    }



@app.put("/api/links/{link_id}", response_model=FavoriteLinkItem)
def update_favorite_link(link_id: int, updated_link: FavoriteLinkCreate) -> dict[str, Any]:
    """Edita um link favorito preservando sua posição na lista."""
    if link_id <= 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID deve ser maior que zero.",
        )

    validated_url = validate_http_url(updated_link.url)
    link_data = model_to_dict(updated_link)
    link_data["title"] = link_data["title"].strip()
    link_data["category"] = link_data["category"].strip()
    link_data["url"] = validated_url
    link_data["note"] = link_data.get("note", "").strip()

    if not link_data["title"] or not link_data["category"]:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome e categoria não podem conter apenas espaços.",
        )

    with LINKS_LOCK:
        links = read_links()

        index = next(
            (i for i, item in enumerate(links) if int(item.get("id", -1)) == link_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Link com ID {link_id} não encontrado.",
            )

        normalized_url = validated_url.rstrip("/").casefold()

        duplicate = any(
            int(item.get("id", -1)) != link_id
            and str(item.get("url", "")).strip().rstrip("/").casefold() == normalized_url
            for item in links
            if isinstance(item, dict)
        )

        if duplicate:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="Já existe outro favorito com esta URL.",
            )

        saved = {"id": link_id, **link_data}
        links[index] = saved
        write_links(links)

    return saved


@app.delete("/api/links/{link_id}")
def delete_favorite_link(link_id: int) -> dict[str, Any]:
    """Exclui um link favorito."""
    if link_id <= 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID deve ser maior que zero.",
        )

    with LINKS_LOCK:
        links = read_links()

        link_to_delete = next(
            (item for item in links if int(item.get("id", -1)) == link_id),
            None,
        )

        if link_to_delete is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Link com ID {link_id} não encontrado.",
            )

        write_links([
            item for item in links
            if int(item.get("id", -1)) != link_id
        ])

    return {
        "message": "Link excluído com sucesso.",
        "deleted_id": link_id,
        "deleted_title": link_to_delete.get("title"),
    }


@app.get("/api/apps", response_model=list[AppItem])
def list_apps() -> list[dict[str, Any]]:
    """Lista todos os aplicativos cadastrados."""
    with DATA_LOCK:
        return read_apps()


@app.post(
    "/api/apps",
    response_model=AppItem,
    status_code=status.HTTP_201_CREATED,
)
def add_app(new_app: AppItem) -> dict[str, Any]:
    """Adiciona um novo aplicativo e impede IDs duplicados."""
    validated_url = validate_http_url(new_app.url)
    app_data = model_to_dict(new_app)
    app_data["url"] = validated_url

    app_data["name"] = app_data["name"].strip()
    app_data["category"] = app_data["category"].strip()
    app_data["icon"] = app_data["icon"].strip()
    app_data["description"] = app_data["description"].strip()

    if not all([
        app_data["name"],
        app_data["category"],
        app_data["icon"],
        app_data["description"],
    ]):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome, categoria, ícone e descrição não podem conter apenas espaços.",
        )

    with DATA_LOCK:
        apps = read_apps()

        if any(int(item.get("id", -1)) == new_app.id for item in apps):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"Já existe um aplicativo com o ID {new_app.id}.",
            )

        apps.append(app_data)
        write_apps(apps)

    return app_data


@app.patch("/api/apps/{app_id}/favorite", response_model=AppItem)
def toggle_app_favorite(app_id: int) -> dict[str, Any]:
    """Alterna o status de favorito de um aplicativo."""
    if app_id <= 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID deve ser maior que zero.",
        )

    with DATA_LOCK:
        apps = read_apps()

        index = next(
            (i for i, item in enumerate(apps) if int(item.get("id", -1)) == app_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Aplicativo com ID {app_id} não encontrado.",
            )

        apps[index]["favorite"] = not bool(apps[index].get("favorite", False))
        updated_app = apps[index]
        write_apps(apps)

    return updated_app


@app.put("/api/apps/order")
def reorder_apps(order: AppsOrderRequest) -> dict[str, Any]:
    """Persiste a ordem escolhida pelo usuário na galeria APPS."""
    if not order.ids:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="A ordem dos aplicativos não pode estar vazia.",
        )

    with DATA_LOCK:
        apps = read_apps()

        current_ids = [int(item.get("id", -1)) for item in apps]
        requested_ids = [int(app_id) for app_id in order.ids]

        if len(requested_ids) != len(set(requested_ids)):
            raise HTTPException(
                status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
                detail="A nova ordem contém IDs duplicados.",
            )

        if set(requested_ids) != set(current_ids) or len(requested_ids) != len(current_ids):
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail="A lista de aplicativos mudou. Atualize a página e tente novamente.",
            )

        apps_by_id = {int(item["id"]): item for item in apps}
        reordered_apps = [apps_by_id[app_id] for app_id in requested_ids]
        write_apps(reordered_apps)

    return {
        "message": "Ordem dos aplicativos salva com sucesso.",
        "ids": requested_ids,
    }



@app.put("/api/apps/{app_id}", response_model=AppItem)
def update_app(app_id: int, updated_app: AppItem) -> dict[str, Any]:
    """Edita um aplicativo existente preservando sua posição no JSON."""
    if app_id <= 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID deve ser maior que zero.",
        )

    if int(updated_app.id) != int(app_id):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID do aplicativo não pode ser alterado durante a edição.",
        )

    validated_url = validate_http_url(updated_app.url)
    app_data = model_to_dict(updated_app)
    app_data["url"] = validated_url
    app_data["name"] = app_data["name"].strip()
    app_data["category"] = app_data["category"].strip()
    app_data["icon"] = app_data["icon"].strip()
    app_data["description"] = app_data["description"].strip()

    if not all([
        app_data["name"],
        app_data["category"],
        app_data["icon"],
        app_data["description"],
    ]):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="Nome, categoria, ícone e descrição não podem conter apenas espaços.",
        )

    old_icon_to_remove: Path | None = None

    with DATA_LOCK:
        apps = read_apps()

        index = next(
            (i for i, item in enumerate(apps) if int(item.get("id", -1)) == app_id),
            None,
        )

        if index is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Aplicativo com ID {app_id} não encontrado.",
            )

        old_icon = str(apps[index].get("icon", ""))
        new_icon = str(app_data.get("icon", ""))

        # Favoritar é uma ação independente do formulário de edição.
        # Preserva o estado atual para evitar que uma edição comum
        # desmarque o aplicativo por acidente.
        app_data["favorite"] = bool(apps[index].get("favorite", False))

        apps[index] = app_data
        write_apps(apps)

        # Se um ícone local foi substituído, remove o arquivo antigo somente
        # depois que o JSON foi salvo com sucesso.
        if (
            old_icon
            and old_icon != new_icon
            and old_icon.startswith("/icons/")
        ):
            old_icon_to_remove = ICONS_DIR / Path(old_icon).name

    if old_icon_to_remove is not None:
        try:
            old_icon_to_remove.unlink(missing_ok=True)
        except OSError:
            # A edição já foi salva; falha na limpeza não deve invalidá-la.
            pass

    return app_data


@app.delete("/api/apps/{app_id}")
def delete_app(app_id: int) -> dict[str, Any]:
    """Exclui um aplicativo pelo ID."""
    if app_id <= 0:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail="O ID deve ser maior que zero.",
        )

    with DATA_LOCK:
        apps = read_apps()

        app_to_delete = next(
            (item for item in apps if int(item.get("id", -1)) == app_id),
            None,
        )

        if app_to_delete is None:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Aplicativo com ID {app_id} não encontrado.",
            )

        updated_apps = [
            item for item in apps
            if int(item.get("id", -1)) != app_id
        ]

        write_apps(updated_apps)

        icon_value = str(app_to_delete.get("icon", ""))
        if icon_value.startswith("/icons/"):
            icon_file = ICONS_DIR / Path(icon_value).name
            try:
                icon_file.unlink(missing_ok=True)
            except OSError:
                # O cadastro já foi removido; falha de limpeza do arquivo não deve impedir a operação.
                pass

    return {
        "message": "Aplicativo excluído com sucesso.",
        "deleted_id": app_id,
        "deleted_name": app_to_delete.get("name"),
    }


# ============================================================
# EXECUÇÃO
# ============================================================

HUB_URL = "http://127.0.0.1:8000"


def hub_port_is_open(host: str = "127.0.0.1", port: int = 8000) -> bool:
    """Evita iniciar uma segunda instância do servidor local."""
    try:
        with socket.create_connection((host, port), timeout=0.35):
            return True
    except OSError:
        return False


def open_hub_in_browser() -> None:
    """Abre o Hub no navegador padrão sem bloquear o servidor."""
    try:
        webbrowser.open(HUB_URL, new=2)
    except Exception:
        pass


def ensure_frozen_stdio() -> None:
    """Cria streams válidos quando o PyInstaller usa --noconsole.

    No modo --noconsole, sys.stdout/sys.stderr podem ser None. O formatter
    padrão do Uvicorn chama stream.isatty(), o que derruba o executável
    antes do servidor iniciar. Redirecionar os streams ausentes para
    os.devnull mantém o aplicativo sem console e compatível com logging.
    """
    if sys.stdin is None:
        sys.stdin = open(
            os.devnull,
            "r",
            encoding="utf-8",
            errors="replace",
        )

    if sys.stdout is None:
        sys.stdout = open(
            os.devnull,
            "w",
            encoding="utf-8",
            errors="replace",
        )

    if sys.stderr is None:
        sys.stderr = open(
            os.devnull,
            "w",
            encoding="utf-8",
            errors="replace",
        )


if __name__ == "__main__":
    frozen_app = bool(getattr(sys, "frozen", False))

    if frozen_app:
        ensure_frozen_stdio()

    if frozen_app and hub_port_is_open():
        open_hub_in_browser()
        raise SystemExit(0)

    if frozen_app:
        browser_timer = threading.Timer(1.25, open_hub_in_browser)
        browser_timer.daemon = True
        browser_timer.start()

    uvicorn.run(
        app,
        host="127.0.0.1",
        port=8000,
        log_level="info",
    )
