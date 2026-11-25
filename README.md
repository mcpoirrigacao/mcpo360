# MCPO 360

App único dos sistemas internos da MCPO Irrigação (MAQCAMPO): um login Microsoft, um menu com os módulos liberados para cada e-mail e uma **Central de acessos** onde se cadastra o e-mail e se marca o nível em cada módulo.

Endereço: https://mcpoirrigacao.github.io/mcpo360/

## Estrutura

| Arquivo | O que é |
|---|---|
| `index.html`, `app.css`, `app.js` | Casca: login, menu, início, Central de acessos, uso dos módulos e sincronização das pastas do SharePoint |
| `redirect.html` | Não usado (o login volta para `/mcpo360/`) |
| `modulos/<id>/index.html` | Cada app antigo, adaptado para rodar dentro do MCPO 360 |

Módulos: `painel`, `assistencia`, `pecas`, `compras`, `orcamento`, `obras`, `entregas`, `crmsam`, `crmcomercial`, `followup`, `treinamentos`.

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

## Entrega Técnica (`entregas`)

- Pasta no SharePoint: `Documentos Compartilhados › Entrega Técnica` (criada pela sincronização da Central, se não existir).
  - `_dados/entregas_dados.json`: entregas, parâmetros e pastas dos clientes (gravação com ETag e mesclagem por registro).
  - `_dados/obras_concluidas.json`: gravado pelo módulo Obras (admin) com as obras concluídas, sem valores. Obra concluída a partir da data de corte (Parâmetros) vira "Pendente ET".
  - `Clientes/<cliente>/<ET-AAAA-NNN - equipamentos - fazenda>/`: anexos enviados pelo módulo (termo, fotos, aprovação da fábrica, comprovante).
- Níveis: Admin (tudo + Parâmetros), Consultor (cria/atualiza entregas e anexos), Comum (só visualiza). Funções do Azure: `Entregas.Admin`, `Entregas.Consultor`, `Entregas.Comum`.
- Admins do Obras recebem escrita na pasta da Entrega Técnica (para gravar `obras_concluidas.json`).
- Etapas: Pendente ET → Em andamento → Enviar à fábrica → Aguardando fábrica → Aguardando pagamento → Paga; divergência é uma marca na entrega.
- Entrega parcial: ao concluir em campo escolhe-se Total, Parcial ou Nada entregue. Parcial e nada entregue exigem o termo de impossibilidade (anexo) + motivo e previsão de retorno. Parcial divide os equipamentos (entregue segue para a fábrica; o pendente vira um lote de retorno com `retornoDe`). A fábrica pode não pagar a parcial (`semPagamento`): ela é quitada junto com o pagamento do retorno.
- Grupos: cliente → obra → lotes. Cada entrega técnica é um lote da obra (`obraId`); sem obra, o grupo é cliente + fazenda ou um grupo escolhido (`grupoId`). Total de pivôs = campo "Pivôs" da obra (Obras; sugerido pela tarefa ELET-000 "Ligação painel e parametrização (por pivô)"), ajustável por grupo em `grupos/<chave>.totalPivos`.
