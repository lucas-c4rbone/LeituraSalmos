# GeradorSalmos

Projeto em Python para gerar automaticamente apresentacoes do Microsoft PowerPoint para leitura biblica em uma igreja.

O programa usa exclusivamente `pywin32` (`win32com.client`) para controlar o PowerPoint instalado no Windows. Ele nao usa `python-pptx`, nao usa VBA e nao recria caixas de texto: apenas duplica slides do modelo e substitui os placeholders existentes, preservando fonte, cor, tamanho, sombra, alinhamento, plano de fundo, animacoes e transicoes.

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

3. Na janela aberta:
   - abra o Busca;
   - copie a leitura biblica;
   - cole o texto no campo `Scripture`;
   - informe o titulo da leitura;
   - clique em `Generate PowerPoint`.

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

Esse atalho tambem gera a apresentacao.

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

## Recursos de producao

- Desabilita os campos durante a geracao.
- Evita cliques duplicados no botao de gerar.
- Mostra progresso animado e status ao vivo.
- Ao terminar, mostra nome do arquivo, quantidade de versiculos e tempo de geracao.
- Abre o Windows Explorer selecionando o arquivo gerado.
- Limpa os campos apos sucesso e foca novamente o campo `Scripture`.
- Lembra a ultima posicao da janela.
- Centraliza a janela no primeiro uso.
- Detecta uma leitura valida na area de transferencia ao iniciar e pergunta se deve colar.
- Usa caminhos relativos ao executavel, adequado para uso com PyInstaller.
