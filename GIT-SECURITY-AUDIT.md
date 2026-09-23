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


## Comandos individuais

A partir da v2.24.2, o Dashboard oferece comandos Git individuais por
ações estruturadas. Não há campo de terminal livre.

Ações que podem colocar conteúdo no histórico ou enviar ao remoto
(`add`, `commit` e `push`) são bloqueadas quando a auditoria encontra
riscos altos.

`git rm` foi implementado somente com `--cached`, portanto remove do
índice do Git sem apagar o arquivo do computador.


## Gerenciador da conta GitHub — v2.25.0

O painel usa o `gh` autenticado e não persiste tokens.
Mudanças para visibilidade pública pedem confirmação explícita.
Rename e exclusão permanente exigem digitação exata de `owner/repo`.


## Política zero-credenciais — v2.25.1

- GitHub: usa a sessão externa do `gh auth`.
- Sites: usam exclusivamente a sessão existente no navegador.
- Aplicativos: usam a sessão do próprio aplicativo.
- Favoritos do navegador: leitura/importação desativada.
- Links pessoais: somente em memória; não são gravados.
- Saídas Git: tokens e credenciais em URLs são mascarados.
- APIs operacionais: `Cache-Control: no-store`.


## Detecção do GitHub CLI — v2.25.2

A localização de `gh.exe` é feita dinamicamente e não é persistida.
O frontend recebe apenas uma identificação genérica da origem da
detecção (`PATH`, `Program Files`, `WinGet Links`, etc.), nunca o
caminho pessoal completo.


## Workspace Explorer — v2.26.0
A raiz operacional padrão é `D:\python\CHATGPT` e não é gravada como
preferência. O antigo `git_projects.json` é removido e caminhos Git
selecionados passam a existir apenas em memória durante a sessão.
O Workspace não exclui nem renomeia pastas físicas.
