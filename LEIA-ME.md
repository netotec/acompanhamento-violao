# Link do aluno no GitHub Pages

Esta pasta é o que vai para o GitHub Pages (`https://netotec.github.io/acompanhamento-violao/`):

- `index.html`: página do aluno. Lê `?t=CÓDIGO` e abre os exercícios.
- `professor.html`: painel do professor (pede a senha como sempre).

Os dois arquivos são **gerados** a partir de `src/Index.html` e `src/Painel.html`. Não edite à mão:

```
node scripts/gerar-pages.js
```

## Como funciona

As páginas são a mesma tela do Apps Script, mas falam com o web app por requisições **sem cookies**
(`fetch` com `credentials: 'omit'`). Assim o Google não troca de conta (`/u/3/...`) nem pede login, e o
Safari não bloqueia nada, porque não há quadro (iframe). O servidor responde em JSON
(`doGet` com `?modo=json` e `doPost`, em `Code.gs`).

## Atualizar o GitHub Pages

1. Rode `node scripts/gerar-pages.js` sempre que mudar `src/Index.html` ou `src/Painel.html`.
2. Publique o Apps Script: `node scripts/publicar.js "descrição"`.
3. Envie `index.html`, `professor.html` e `.nojekyll` para a raiz do repositório
   `netotec/acompanhamento-violao` (branch `main`).

## Configuração (uma vez)

- Em **Settings > Pages**: *Deploy from a branch*, branch `main`, pasta `/ (root)`.
- No painel: **Enviar link**, campo "Endereço dos links" com `https://netotec.github.io/acompanhamento-violao/`.

## Observações

- O endereço do web app (`EXEC`) está em `scripts/gerar-pages.js`. Se criar outra implantação, troque lá e gere de novo.
- O código do aparelho do aluno fica no armazenamento da própria página do GitHub Pages. Quem já tinha
  aparelho cadastrado pelo link antigo conta como aparelho novo (limite de 2). Se ficar bloqueado, use
  **Alunos > Editar > Liberar aparelho novo**.
- O Safari apaga dados de sites não visitados há 7 dias; para quem usa uma vez por semana, o código do
  aparelho pode ser recriado. Se isso incomodar, dá para afrouxar a trava de aparelho.
