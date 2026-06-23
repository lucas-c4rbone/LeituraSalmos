# CHANGELOG

Todas as mudanças relevantes deste projeto devem ser documentadas neste arquivo.

Este changelog segue o formato inspirado em Keep a Changelog e Semantic Versioning.

## [Unreleased]

### Added

- Documentação de produção criada:
  - `README.md` atualizado para operação e release
  - `INSTALL.md`
  - `BUILD.md`
  - `CHANGELOG.md`

### Changed

- `script.iss` profissionalizado com:
  - `AppId` estável
  - metadados de versão do instalador
  - diretório por usuário (`%LOCALAPPDATA%\\Programs`)
  - arquitetura `x64compatible`
  - setup logging e melhorias de UX

## [2.0.0] - 2026-06-23

### Added

- Melhorias de estabilidade e qualidade:
  - `ui/__init__.py` incluído para consistência de módulo
  - cache de lookup de placeholders no gerador para reduzir varreduras repetidas

### Changed

- Logging da aplicação ajustado para evitar efeitos colaterais no root logger, mantendo logs em arquivo.

### Fixed

- Falha estrutural de testes relacionada à ausência de `ui/__init__.py`.

## [1.1.0] - 2026-06-23

### Added

- Instalador via Inno Setup com atalho opcional em desktop e execução pós-instalação.

### Notes

- Tag histórica baseada na versão de instalador (`script.iss`).
