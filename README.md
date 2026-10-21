# 🎢 Roteiro Orlando 2027

Sistema de votação em grupo para montar o roteiro da viagem Tampa → Orlando → Miami → Tampa (08 a 19/02/2027).

Cada pessoa abre o link no celular, escolhe o nome, cria um PIN e vota nos parques, atrações e passeios. O organizador trava os dias com ingresso comprado e gera o roteiro automaticamente pelos votos.

## Arquivos

| Arquivo | O que é |
|---|---|
| `Codigo.gs` | Backend (Google Apps Script). A planilha vinculada é o banco de dados. |
| `Index.html` | Tela do app (mobile). |

## Como instalar (uma vez)

1. Crie uma planilha nova no Google Sheets chamada **Roteiro Orlando 2027**.
2. Menu **Extensões → Apps Script**.
3. Apague o conteúdo de `Código.gs` e cole todo o `Codigo.gs`.
4. Clique em **+ → HTML**, nomeie **Index** e cole todo o `Index.html`.
5. Salve, escolha a função **setup** e clique em **Executar**. Autorize o acesso.
6. Volte à planilha: as abas foram criadas com o catálogo real. O **PIN do organizador** fica na aba `Config` (linha `pin_admin`).
7. No Apps Script: **Implantar → Nova implantação → Tipo: App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**
8. Copie o link do app e mande no grupo da família.

## Abas da planilha

- **Config** — nome da viagem, datas e PIN do organizador
- **Dias** — os 11 dias, cidade, tipo, local e se está fixo (🔒)
- **Locais** — 18 parques/passeios (Disney, Universal, Epic Universe, SeaWorld, Busch Gardens, LEGOLAND, aquáticos, NASA, Miami, Tampa, Clearwater, compras)
- **Atracoes** — 194 atrações por área, categoria e altura mínima. Para esconder uma, coloque `NAO` na coluna `ativo`.
- **Pessoas**, **Votos**, **Roteiro** — preenchidas pelo app

## Como o roteiro é montado

- Votos: cada pessoa vale 1. 😍 Quero muito e 🙂 Tanto faz contam a favor, 🙅 Não quero conta contra. Saldo = a favor − contra; empate → mais 😍; entre atrações empatadas, a mais perto da anterior no parque vem primeiro.
- Dias **fixos** mantêm o local escolhido (ingressos já comprados).
- Os outros dias de parque recebem o local com mais pontos da região, sem repetir enquanto houver opção. Locais com saldo negativo ficam de fora.
- Dias em Orlando podem receber a NASA (Kennedy Space Center); o dia de retorno escolhe entre os locais de Tampa.
- Dentro de cada parque entram as atrações com pontos positivos, agrupadas por área, começando pela área da atração mais votada.
- Atrações com altura mínima mostram quem do grupo não pode ir.

## Fotos

As fotos vêm do Wikimedia Commons (licença livre) e ficam salvas na planilha depois da primeira busca. Para trocar uma foto, cole o link de outra imagem na coluna `foto_url`.

> Alturas e atrações conferidas para a situação de 2026/2027 (DinoLand e Rivers of America já fechados; Muppets no Rock 'n' Roller Coaster; Epic Universe aberto). Confira sempre no app oficial do parque perto da viagem.
