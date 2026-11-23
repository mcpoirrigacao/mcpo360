# MCPO 360

App único dos sistemas internos da MCPO Irrigação (MAQCAMPO): um login Microsoft, um menu com os módulos liberados para cada e-mail e uma **Central de acessos** onde se cadastra o e-mail e se marca o nível em cada módulo.

Endereço: https://mcpoirrigacao.github.io/mcpo360/

## Estrutura

| Arquivo | O que é |
|---|---|
| `index.html`, `app.css`, `app.js` | Casca: login, menu, início, Central de acessos, uso dos módulos e sincronização das pastas do SharePoint |
| `redirect.html` | Não usado (o login volta para `/mcpo360/`) |
| `modulos/<id>/index.html` | Cada app antigo, adaptado para rodar dentro do MCPO 360 |

Módulos: `painel`, `assistencia`, `pecas`, `compras`, `obras`, `orcamento`, `crmsam`, `crmcomercial`, `followup`, `treinamentos`.

## Acessos

- Arquivo único: SharePoint MCPOIRRIGACAO › Documentos Compartilhados › `MCPO360/acessos.json` (`usuarios[email].modulos[modulo] = nível`).
- Funções do app no Azure (registro "MCPO API Github") valem por cima: `MCPO.Admin` = admin em tudo e na Central; `Pecas.Admin`, `CRMSAM.Editor` etc. valem no módulo.
- Ao salvar na Central, o MCPO 360 libera/retira a permissão nas pastas de cada módulo (registro do que liberou em `_mcpo_permissoes.json` em cada pasta). Retiradas em massa pedem confirmação.
- Os módulos perguntam o nível à casca (`window.top.MCPO360.nivel('<id>')`). Aberto fora do MCPO 360, o módulo volta para `/mcpo360/#<id>` (exceto links de confirmação/avaliação do Treinamentos).
- FollowUP: nível "Equipe" — o que a pessoa preenche/vê continua definido nas equipes dentro do FollowUP.
- CRM Comercial mantém o login próprio da planilha.

## Uso

Cada pessoa grava um contador por módulo/dia em `MCPO360/uso/<email>.json`. A tela "Uso dos módulos" (admins da Central) mostra os últimos 30 dias.

## Pré-requisitos no Azure

No registro **MCPO API Github** › Autenticação › Aplicativo de página única, cadastrar:
`https://mcpoirrigacao.github.io/mcpo360/` (já cadastrado)
