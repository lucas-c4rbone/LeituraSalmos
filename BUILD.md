# BUILD

Guia de build de produção do executável e do instalador.

## Requisitos de Build

- Windows 10/11
- Python 3.13
- Inno Setup 6 (ISCC)
- PowerShell

## Preparação do Ambiente

Na raiz do projeto:

```powershell
py -m pip install -r requirements.txt
py -m pip install pyinstaller
```

## 1) Executar Testes (Recomendado)

```powershell
py -m pytest -q
```

## 2) Gerar Executável (.exe)

Build oficial usando spec:

```powershell
py -m PyInstaller --noconfirm --clean "Leitura Bíblica.spec"
```

Saída esperada:

- `dist\\Leitura Bíblica.exe`

Observações:

- O arquivo `.spec` já inclui `modelo.pptx` em `datas`.
- O executável é gerado em modo janela (`console=False`).

## 3) Gerar Instalador (.exe)

Com Inno Setup instalado, execute:

```powershell
"C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe" "script.iss"
```

Saída esperada:

- `Output\\Setup_Leitura_Biblica_v<versao>.exe`

## 4) Pipeline Manual Completo

```powershell
py -m pytest -q
py -m PyInstaller --noconfirm --clean "Leitura Bíblica.spec"
"C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe" "script.iss"
```

## Versionamento de Release

- Atualize versão do app em `ui/theme.py` (`APP_VERSION`).
- Atualize versão do instalador em `script.iss` (`MyAppVersion`).
- Mantenha as duas versões sincronizadas antes da release.

## Checklist de Produção

- Testes passando.
- `dist\\Leitura Bíblica.exe` abre sem traceback.
- Geração de apresentação funciona com `modelo.pptx` válido.
- Instalador cria atalhos e inicia o app no fim.
- Desinstalação remove binários instalados.
