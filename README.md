# GeradorSalmos

Gerador profissional de apresentações do Microsoft PowerPoint para leitura bíblica em igrejas, com interface moderna Fluent Design e animações suaves.

O programa usa exclusivamente `pywin32` (`win32com.client`) para controlar o PowerPoint instalado no Windows. Ele não usa `python-pptx`, não usa VBA e não recria caixas de texto: apenas duplica slides do modelo e substitui os placeholders existentes, preservando fonte, cor, tamanho, sombra, alinhamento, plano de fundo, animações e transições.

## Características

- **Interface Fluent Design** — Tema escuro moderno com animações suaves (Fluent Reveal, entrance animations, spring easing)
- **Splash screen** — Tela de carregamento elegante enquanto a UI é inicializada
- **Progresso visual** — Barra de progresso animada com status em tempo real
- **Intervalo de versículos** — Seção expansível para gerar apenas um trecho específico
- **Responde em tempo real** — Detecção automática de leitura válida na área de transferência ao iniciar
- **Persistência** — Lembra posição da janela entre sessões
- **Tratamento robusto de erros** — Mensagens amigáveis, nunca expõe tracebacks
- **Logs detalhados** — Histórico completo de operações em `logs.log`

## Estrutura

```text
GeradorSalmos/
├── main.py
├── modelo.pptx
├── requirements.txt
└── README.md
```

## Requisitos

- Windows
- Python 3.13
- Microsoft PowerPoint instalado
- Dependencias do arquivo `requirements.txt`

Instale as dependencias:

```powershell
py -m pip install -r requirements.txt
```

## Como usar

1. Coloque o arquivo `modelo.pptx` na mesma pasta do `main.py`.
2. Execute:

```powershell
py main.py
```

3. Na janela aberta (após a splash screen):
   - Abra o Busca.com.br;
   - Copie a leitura bíblica (ex: Salmos 32:1-5);
   - Cole o texto no campo `Scripture`;
   - Informe o título da leitura;
   - (Opcional) Use **"▶ Intervalo de versículos"** para gerar apenas versículos específicos;
   - Clique em **`▶ Gerar PowerPoint (Ctrl+Enter)`**.

O arquivo final sera salvo automaticamente na mesma pasta do programa, com o formato:

```text
<Livro> <Capitulo>.pptx
```

Exemplos:

```text
Salmos 32.pptx
Proverbios 3.pptx
Isaias 53.pptx
Mateus 5.pptx
```

Apos gerar, o programa mostra uma mensagem de sucesso com o nome do arquivo e a pasta de saida, e abre essa pasta automaticamente no Windows Explorer.
Se o arquivo final ja existir, o programa pergunta antes de substituir.

Atalho:

```text
Ctrl+Enter
```

Este atalho também gera a apresentação (sem digitar o título).

### Intervalo de versículos

Expanda a seção **"▶ Intervalo de versículos (opcional)"** para gerar apenas um trecho:

- **De**: Número do primeiro versículo (ex: 1)
- **Até**: Número do último versículo (ex: 5)
- Deixe em branco para gerar todos

### Atalho

## Formato do texto colado

Cole no campo `Scripture` o texto copiado do Busca neste padrao:

```text
1 BEM-AVENTURADO aquele cuja transgressao e perdoada...
2 Bem-aventurado...
3 ...
Livro: Salmos
Capitulo: 32
```

O programa valida automaticamente:

- pelo menos um versiculo numerado;
- linha `Livro:`;
- linha `Capitulo:`.

Tambem aceita `Capítulo:` com acento.

## Placeholders do modelo

O `modelo.pptx` deve conter 4 slides:

1. Slide de titulo com:
   - `{livro_cap}` ou `livro cap`
   - `{titulo}` ou `titulo`
2. Slide modelo do pastor, com texto branco:
   - `{versiculo}` ou o placeholder equivalente do modelo
3. Slide modelo da igreja, com texto amarelo:
   - `{versiculo}` ou o placeholder equivalente do modelo
4. Slide modelo final, com texto amarelo sublinhado:
   - `{versiculo}` ou o placeholder equivalente do modelo

## Regras de geracao

- Slide 1 recebe `<Livro> <Capitulo>` e o titulo informado na interface.
- O ultimo versiculo sempre usa o slide sublinhado.
- Os demais versiculos alternam automaticamente:
  - pastor/branco;
  - igreja/amarelo;
  - pastor/branco;
  - igreja/amarelo.
- As caixas de texto dos versiculos sao centralizadas no slide.
- O texto dos versiculos permanece justificado.
- A quantidade de slides e gerada automaticamente.

## Tratamento de erros

A interface mostra mensagens amigaveis para problemas como:

- texto colado vazio ou invalido;
- falta da linha `Livro:`;
- falta da linha `Capitulo:`;
- ausencia de versiculos numerados;
- `modelo.pptx` inexistente;
- PowerPoint nao instalado;
- placeholders ausentes no modelo;
- erro de permissao ao salvar;
- falhas retornadas pelo PowerPoint.

O usuario nunca recebe tracebacks Python na tela. Detalhes tecnicos, historico de geracao e erros sao registrados em:

```text
logs.log
```

## Recursos de Produção

- ✅ **Splash screen** — Tela elegante de carregamento (500ms mínimo)
- ✅ **Animações Fluent** — Entrance animada (header → card → tiles), scale-in, hover effects
- ✅ **Desabilita campos durante geração** — Previne cliques duplicados
- ✅ **Progresso visual** — Barra animada + mensagens de status em tempo real
- ✅ **Ao terminar** — Nome do arquivo, quantidade de versículos e tempo de geração
- ✅ **Abre Explorer** — Automático com arquivo pré-selecionado
- ✅ **Limpa campos** — Após sucesso, foca novamente em `Scripture`
- ✅ **Detecta clipboard** — Pergunta se deve colar leitura válida ao iniciar
- ✅ **Lembra geometria** — Última posição da janela persiste
- ✅ **Caminhos relativos** — Compatível com PyInstaller
- ✅ **Temas escuros** — Tema Fluent Blue com contraste otimizado (WCAG AA)
- ✅ **Tratamento de erros** — Mensagens amigáveis, nunca expõe tracebacks Python
