# INSTALL

Guia de instalação para usuários finais e operadores em Windows.

## Escopo

Este documento cobre:

- pré-requisitos
- instalação via instalador `.exe`
- execução em modo desenvolvimento (Python)
- desinstalação
- solução de problemas comuns

## Pré-requisitos

- Windows 10/11 (64 bits recomendado)
- Microsoft PowerPoint instalado e funcional
- Para modo desenvolvimento: Python 3.13 e `pip`

Dependências Python utilizadas:

- `pywin32>=306`
- `customtkinter>=5.2.0`

## Instalação para Usuário Final (Recomendada)

1. Obtenha o instalador em `Output/` com nome no padrão:
   - `Setup_Leitura_Biblica_v<versao>.exe`
2. Execute o instalador.
3. Siga o assistente de instalação.
4. (Opcional) Marque a criação do atalho na área de trabalho.
5. Conclua e abra o aplicativo.

Diretório padrão de instalação por usuário:

- `%LOCALAPPDATA%\\Programs\\Leitura Bíblica`

## Instalação para Desenvolvimento (Python)

1. Abra o terminal na raiz do projeto.
2. Instale dependências:

```powershell
py -m pip install -r requirements.txt
```

3. Execute o app:

```powershell
py main.py
```

## Arquivos Necessários em Runtime

- `modelo.pptx` deve estar disponível para geração.
- `logs.log` é criado automaticamente em:
  - `%LOCALAPPDATA%\\LeituraBiblica\\logs.log`

## Desinstalação

- Use "Aplicativos Instalados" do Windows e remova `Leitura Bíblica`.
- O desinstalador remove os arquivos instalados em `%LOCALAPPDATA%\\Programs\\Leitura Bíblica`.
- Dados de usuário (como logs/config em `%LOCALAPPDATA%\\LeituraBiblica`) podem permanecer para diagnóstico.

## Troubleshooting

### Erro: PowerPoint não abre

- Verifique se o Microsoft PowerPoint está instalado e abre manualmente.
- Repare a instalação do Office, se necessário.

### Erro: `modelo.pptx` não encontrado

- Garanta que o arquivo `modelo.pptx` esteja no local esperado pelo aplicativo.

### Erro de permissão ao salvar `.pptx`

- Feche o arquivo de saída no PowerPoint.
- Verifique se a pasta de saída permite escrita.

### App não inicia após instalação

- Execute o executável diretamente de `%LOCALAPPDATA%\\Programs\\Leitura Bíblica`.
- Verifique logs em `%LOCALAPPDATA%\\LeituraBiblica\\logs.log`.
