TECH TOOL HUB v2.23.1 - BUILD ENXUTO
======================================

CORREÇÃO
--------
O build anterior chegou à fase:

    Performing binary vs. data reclassification (9683 entries)

porque o PyInstaller estava incluindo bibliotecas opcionais instaladas
no Python global que não são usadas pelo TECH TOOL HUB.

A v2.23.1:
- remove --collect-all uvicorn;
- mantém apenas os hidden imports necessários do Uvicorn;
- exclui IPython, NumPy, Matplotlib, PyQt, pygame, pytest, Sphinx,
  Jupyter e outros pacotes não usados pelo Hub.

COMO CONTINUAR
--------------
Se o build antigo ainda estiver parado, pressione:

    CTRL + C

Substitua app.py e build-msix.ps1 pelos desta versão.

Depois:

    Set-ExecutionPolicy -Scope Process Bypass
    .\build-msix.ps1 -Install

A geração do EXE pelo botão "Criar Executável" do Dashboard também
usa agora o mesmo modo enxuto.

SAÍDA MSIX
----------
    release\TechToolHub_2.23.1.0_x64.msix

SAÍDA EXE PELO DASHBOARD
------------------------
    release\exe\TechToolHub\TechToolHub.exe

DADOS
-----
    %LOCALAPPDATA%\TechToolHub


HOTFIX DO MANIFESTO MSIX
------------------------
Se o MakeAppx informar:

    The DefaultTile element must specify the Wide310x150Logo attribute
    if the Square310x310Logo attribute is specified.

esta versão já corrige o manifesto e inclui:

    packaging\Assets\Wide310x150Logo.png

COMO RETOMAR SEM REPETIR O PYINSTALLER
--------------------------------------
Se o build anterior já terminou a etapa do PyInstaller e falhou apenas
no MakeAppx, use:

    .\build-msix.ps1 -Install -ResumePackage

A opção -ResumePackage reutiliza:

    build-msix\dist\TechToolHub\TechToolHub.exe

e repete somente a montagem do layout, criação, assinatura e instalação
do MSIX.

Se essa pasta não existir, execute normalmente:

    .\build-msix.ps1 -Install


CORREÇÃO 0x800B0109 - v2.23.2
------------------------------
A versão anterior confiava no certificado em:

    CurrentUser\TrustedPeople

Para instalação de MSIX com certificado autoassinado, esta versão
passa a confiar o certificado em:

    LocalMachine\TrustedPeople

Isso exige permissão de Administrador.

SE O MSIX JÁ FOI GERADO
-----------------------
Não gere novamente. Execute apenas o instalador corrigido como abaixo:

    .\install-msix.ps1 `
        -MsixPath ".\release\TechToolHub_2.23.1.0_x64.msix" `
        -CertificatePath ".\release\TechToolHub.cer"

O script solicita elevação de Administrador automaticamente.

PARA NOVOS BUILDS
-----------------
Use normalmente:

    .\build-msix.ps1 -Install

A instalação passará pelo instalador elevado automaticamente.


CORREÇÃO DE INICIALIZAÇÃO - v2.23.3
-----------------------------------
Erro corrigido:

    Failed to execute script 'pyi_rth_pkgres'
    ModuleNotFoundError: No module named 'jaraco.text'

A versão v2.23.3:
- atualiza o PyInstaller para 6.22.3 antes do build;
- valida/instala jaraco.text, jaraco.functools e jaraco.context;
- inclui explicitamente o namespace jaraco no executável;
- inclui os dados de jaraco.text;
- mantém o modo de build enxuto.

É NECESSÁRIO RECOMPILAR
-----------------------
O erro está dentro do TechToolHub.exe já empacotado.

Execute:

    Set-ExecutionPolicy -Scope Process Bypass
    .\build-msix.ps1 -Install

NÃO use:

    -ResumePackage

nesta correção, pois isso reutilizaria o EXE antigo.

Saída:

    release\TechToolHub_2.23.3.0_x64.msix


CORREÇÃO POWERSHELL/JARACO - v2.23.4
------------------------------------
Problema:
O teste inicial de importação de jaraco retornava traceback em stderr.
Como o script usa:

    $ErrorActionPreference = "Stop"

o Windows PowerShell encerrava o build antes de chegar ao bloco que
instalaria as dependências.

A v2.23.4 permite que o teste falhe normalmente, captura
$LASTEXITCODE, instala jaraco e só então faz a validação final.

Para continuar:

    .\build-msix.ps1 -Install

Não é necessário usar -ResumePackage porque o build parou antes da
geração do novo EXE.


CORREÇÃO UVICORN / --NOCONSOLE - v2.23.5
-----------------------------------------
Erro corrigido:

    AttributeError: 'NoneType' object has no attribute 'isatty'
    ValueError: Unable to configure formatter 'default'

Causa:
O executável é gerado pelo PyInstaller com --noconsole. Nesse modo,
sys.stdout e sys.stderr podem ser None. O formatter padrão do Uvicorn
consulta stream.isatty() durante a inicialização.

Correção:
A v2.23.5 cria streams válidos apontando para os.devnull quando o Hub
está empacotado. Assim:
- o aplicativo continua sem janela de console;
- o Uvicorn consegue configurar o logging;
- stdout/stderr não aparecem para o usuário.

É NECESSÁRIO RECOMPILAR
-----------------------
Execute:

    Set-ExecutionPolicy -Scope Process Bypass
    .\build-msix.ps1 -Install

NÃO use -ResumePackage, porque o EXE anterior contém o problema.

Novo pacote:

    release\TechToolHub_2.23.5.0_x64.msix


DASHBOARD GIT & GITHUB - v2.24.0
--------------------------------
Nova área:

    http://127.0.0.1:8000/git-github

Funções:
1. Checar vazamentos antes de publicar.
2. Automatizar git add / commit / push.
3. Criar GitHub Release com o MSIX mais recente.

Segurança:
- as ações só aceitam acesso local;
- não há campo para token;
- a Release usa a sessão autenticada do GitHub CLI (gh);
- publicação é bloqueada se a auditoria encontrar risco alto.

GitHub CLI:
    winget install --id GitHub.cli
    gh auth login
    gh auth status

Para operar Git/GitHub:
    python app.py

O modo EXE/MSIX instalado não executa publicação do repositório fonte.


SELETOR MULTIPROJETO GIT/GITHUB - v2.24.1
------------------------------------------
O Pipeline local seguro não fica mais preso ao repositório HubTools.

Na tela GIT & GITHUB:
- use o select "Projeto ativo" para trocar entre projetos recentes;
- use "Selecionar pasta" para abrir o seletor nativo do Windows;
- a pasta escolhida fica salva em:
  %LOCALAPPDATA%\TechToolHub\git_projects.json
- o arquivo acima não é criado dentro do repositório.

Todas as ações passam a usar o projeto ativo:
1. checagem de vazamentos;
2. git add / commit;
3. git fetch;
4. git pull --rebase quando o remoto estiver à frente;
5. git push;
6. criação de GitHub Release.

Se o rebase gerar conflito, o Hub aborta o rebase automaticamente e
informa os arquivos conflitantes, sem forçar push.

RELEASES GENÉRICAS
------------------
A tela também lista artefatos existentes em:
    <projeto>\release

Tipos reconhecidos incluem MSIX, EXE, MSI, ZIP, 7Z, WHL e APPX.
Arquivos privados como PFX, PEM e KEY nunca são oferecidos para Release.

O GitHub CLI continua sendo usado com a sessão já autenticada:
    gh auth status


COMANDOS INDIVIDUAIS GIT - v2.24.2
-----------------------------------
A área GIT & GITHUB agora possui um painel de comandos isolados.

Comandos disponibilizados:
- git status
- git init
- git clone
- git add
- git commit
- git switch
- git branch (listar / criar / excluir)
- git fetch origin
- git pull --rebase
- git push
- git log
- git diff
- git remote -v
- git rm -r --cached

Proteções:
- não existe terminal/shell arbitrário no Dashboard;
- Add, Commit e Push passam pela auditoria de vazamentos;
- clone rejeita URLs HTTPS com credenciais/tokens embutidos;
- pull --rebase com conflito é abortado automaticamente;
- exclusão de branch exige confirmação;
- remoção de arquivos usa --cached, preservando os arquivos locais;
- comandos destrutivos como reset --hard não são expostos.

Versão MSIX:
    2.24.2.0


GERENCIADOR DE REPOSITÓRIOS GITHUB - v2.25.0
----------------------------------------------
A área GIT & GITHUB agora lista e administra os repositórios da conta
autenticada pelo GitHub CLI.

Inclui:
- total/públicos/privados/locais/arquivados/forks;
- busca, filtros e ordenação;
- abrir no GitHub e copiar URL;
- detectar clone local conhecido pelo Hub;
- selecionar clone local como projeto ativo;
- clonar repositório;
- criar repositório público/privado e opcionalmente clonar;
- editar descrição, homepage, visibilidade, Issues, Projects, Wiki e
  branch padrão;
- arquivar/restaurar;
- renomear com confirmação e atualização automática do origin local
  quando possível;
- exclusão permanente com dupla confirmação.

Segurança:
- não armazena token GitHub;
- usa a sessão do gh auth;
- mudança para público exige confirmação explícita;
- rename/delete exigem o nome completo owner/repo;
- delete pode exigir o escopo delete_repo no GitHub CLI.

Versão MSIX:
    2.25.0.0


POLÍTICA ZERO-CREDENCIAIS - v2.25.1
------------------------------------
O TECH TOOL HUB não armazena:
- senhas;
- tokens;
- cookies;
- sessões;
- chaves de acesso;
- links pessoais persistentes;
- credenciais embutidas em URLs.

Acesso a apps e sites:
- o Hub apenas abre a URL/aplicativo;
- login/autenticação pertencem ao navegador/site/aplicativo;
- se não houver sessão ativa, o próprio serviço solicitará login.

Links:
- links adicionados manualmente são temporários, apenas em memória;
- links_data.json deixou de ser usado e é removido no startup;
- leitura/importação de favoritos do navegador foi desativada;
- os links temporários desaparecem ao encerrar o processo.

Git/GitHub:
- autenticação usa somente a sessão externa do Git/GitHub CLI;
- nenhum token GitHub é persistido pelo Hub;
- saídas apresentadas no frontend passam por mascaramento de tokens;
- APIs operacionais usam Cache-Control: no-store.

Versão MSIX:
    2.25.1.0


CORREÇÃO DE DETECÇÃO DO GITHUB CLI - v2.25.2
----------------------------------------------
O TECH TOOL HUB não depende mais somente do PATH do processo para
encontrar o GitHub CLI.

No Windows, a detecção procura nesta ordem:
1. PATH do processo;
2. Program Files\GitHub CLI\gh.exe;
3. LocalAppData\Programs\GitHub CLI\gh.exe;
4. LocalAppData\Microsoft\WinGet\Links\gh.exe;
5. pacotes do WinGet em LocalAppData\Microsoft\WinGet\Packages.

Motivo:
um EXE/MSIX pode ter um PATH diferente do PowerShell, especialmente
quando o GitHub CLI foi instalado depois que o aplicativo já estava
aberto.

O caminho completo de gh.exe NÃO é persistido nem enviado ao frontend.
O Dashboard mostra somente a origem genérica da detecção.

Diagnóstico:
    gh --version
    gh auth status
    where.exe gh
    Get-Command gh | Select-Object Source

Versão MSIX:
    2.25.2.0


WORKSPACE EXPLORER - v2.26.0
-----------------------------
O Workspace agora lista as pastas reais de D:\python\CHATGPT em uma
interface inspirada no Explorer do Windows.

Inclui:
- busca instantânea;
- filtros por Em andamento, Pausado, Concluído, Sem cadastro e Git;
- colunas Nome, Status, Progresso, Git e Modificado;
- abrir pasta no Explorer;
- abrir Terminal na pasta;
- selecionar pasta no painel Git & GitHub;
- cadastrar uma pasta ainda sem metadados no Hub;
- abrir o dashboard existente para projeto já cadastrado.

O vínculo com o cadastro é pelo nome da pasta/projeto.
Excluir o cadastro no Hub não apaga a pasta física.
O Workspace não oferece exclusão nem renomeação de diretórios.

Privacidade:
- a raiz operacional não é persistida como preferência;
- git_projects.json legado é removido;
- caminhos Git selecionados ficam somente na memória da sessão.

Versão MSIX: 2.26.0.0


REORGANIZAÇÃO GIT & GITHUB - v2.26.1
--------------------------------------
Os controles de manutenção:
- Criar Executável
- Readme

foram removidos do cabeçalho do Dashboard principal e transferidos para
a página GIT & GITHUB, ao lado de "Atualizar status".

As APIs e o comportamento do build/README foram preservados.

Versão MSIX:
    2.26.1.0


CORREÇÃO DE ARTEFATO DA RELEASE - v2.26.2
------------------------------------------
Problema corrigido:
o painel de Release varria recursivamente a pasta release e podia
selecionar base_library.zip, um componente interno do PyInstaller,
como artefato principal.

Agora:
- base_library.zip nunca é considerado artefato publicável;
- arquivos dentro de _internal são ignorados;
- TechToolHub.exe é priorizado automaticamente;
- para o TECH TOOL HUB, o executável esperado é:
  release\exe\TechToolHub\TechToolHub.exe
- a lista de artefatos marca executáveis com "EXE •".

Observação:
o build atual do botão "Criar Executável" usa PyInstaller --onedir.
Portanto TechToolHub.exe pertence ao conjunto de arquivos da pasta
release\exe\TechToolHub. Para distribuição portátil em outra máquina,
a pasta inteira deve acompanhar o executável ou ser empacotada em ZIP.

Versão MSIX:
    2.26.2.0
