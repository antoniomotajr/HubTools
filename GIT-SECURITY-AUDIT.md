# TECH TOOL HUB — Git & GitHub

O dashboard `/git-github` centraliza:

1. **Checagem de vazamentos**
   - arquivos de dados locais;
   - `.env`, chaves e certificados privados;
   - MSIX/EXE/builds indevidamente rastreados;
   - padrões comuns de credenciais hardcoded.

2. **Publicação Git**
   - auditoria automática antes do staging;
   - `git add -A`;
   - commit;
   - `git push origin <branch atual>`.

3. **GitHub Release**
   - usa `gh` já autenticado;
   - não recebe nem grava token no TECH TOOL HUB;
   - publica o MSIX mais recente encontrado em `release/`;
   - opcionalmente anexa `TechToolHub.cer`.

## Requisitos

```powershell
git --version
gh --version
gh auth status
```

Caso `gh` não esteja instalado:

```powershell
winget install --id GitHub.cli
gh auth login
```

As ações Git/GitHub são liberadas quando o Hub é executado na pasta-fonte:

```powershell
python app.py
```

No EXE/MSIX instalado, essas ações permanecem desabilitadas para evitar operar sobre uma pasta que não é o repositório fonte.


## v2.24.1 — múltiplos projetos

O dashboard Git/GitHub passou a trabalhar com uma lista local de projetos.
A seleção é persistida fora dos repositórios em:

`%LOCALAPPDATA%\TechToolHub\git_projects.json`

A checagem, commit, sincronização, push e Release usam exclusivamente o
repositório do projeto selecionado.

Antes do push, o fluxo executa `git fetch origin` e, quando o remoto está
à frente, `git pull --rebase origin <branch>`. Se houver conflito, o rebase
é abortado para preservar o estado anterior e o push não é realizado.
