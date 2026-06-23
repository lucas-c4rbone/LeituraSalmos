# Gerador de Leitura Bíblica

Aplicativo Windows para gerar apresentações do Microsoft PowerPoint a partir de texto bíblico formatado.

A geração usa automação COM via `pywin32` (`win32com.client`) e preserva o layout do modelo (fontes, cores, alinhamento, animações e transições), substituindo apenas placeholders.

## Recursos Principais

- Interface desktop moderna com `customtkinter`
- Splash screen e feedback visual de progresso
- Geração por intervalo de versículos (opcional)
- Validação amigável de entrada
- Logs de operação para suporte
- Compatível com build via PyInstaller e instalador Inno Setup

## Requisitos

- Windows 10/11
- Microsoft PowerPoint instalado
- Python 3.13 (para execução em desenvolvimento)

Dependências:

- `pywin32>=306`
- `customtkinter>=5.2.0`

## Execução em Desenvolvimento

1. Instale dependências:

```powershell
py -m pip install -r requirements.txt
```

2. Execute o app:

```powershell
py main.py
```

## Uso

1. Abra o app.
2. Cole o texto bíblico no campo principal.
3. Informe o título da leitura.
4. Opcionalmente, defina intervalo (De/Até).
5. Clique em `Gerar PowerPoint`.

Formato esperado da entrada (exemplo):

```text
Livro: Salmos
Capítulo: 23
1 O Senhor é o meu pastor
2 Nada me faltará
```

## Modelo de PowerPoint

O arquivo `modelo.pptx` deve estar disponível para a aplicação e conter placeholders suportados.

Principais placeholders:

- Livro/Capítulo: `{livro_cap}`, `livro cap`, `livro_cap`
- Título: `{titulo}`, `titulo`, `título`
- Versículo: `{versiculo}` e variantes do modelo

## Estrutura do Projeto

```text
LeituraSalmos/
├── app.py
├── main.py
├── requirements.txt
├── Leitura Bíblica.spec
├── script.iss
├── core/
├── ui/
└── tests/
```

## Build e Distribuição

- Gerar executável: consulte `BUILD.md`
- Gerar instalador: consulte `BUILD.md`
- Instalação em cliente final: consulte `INSTALL.md`

## Logs e Diagnóstico

Arquivo de log da aplicação:

- `%LOCALAPPDATA%\LeituraBiblica\logs.log`

## Testes

```powershell
py -m pytest -q
```

## Documentação Relacionada

- `INSTALL.md`
- `BUILD.md`
- `CHANGELOG.md`
