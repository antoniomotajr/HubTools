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
