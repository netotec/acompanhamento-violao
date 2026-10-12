# Link do aluno no GitHub Pages

Esta pasta é o que vai para o GitHub Pages. Ela só tem duas páginas que abrem o web app do Apps Script
dentro de um quadro, num endereço fixo (`https://netotec.github.io/acompanhamento-violao/`):

- `index.html`: página do aluno. Lê `?t=CÓDIGO` e abre os exercícios.
- `professor.html`: abre o painel do professor (pede a senha como sempre).

Por que usar: o link do Google (`script.google.com/...`) pode pedir login ou dar erro quando a pessoa
tem várias contas Google no navegador. No endereço do GitHub Pages abre em qualquer conta ou navegador.

## Publicar (uma vez)

1. No GitHub, crie um repositório **público** chamado `acompanhamento-violao`. Use um repositório
   separado do projeto, só com o conteúdo desta pasta: assim o código do projeto continua privado.
2. Envie `index.html`, `professor.html` e `.nojekyll` para a raiz do repositório.
3. Em **Settings > Pages**, escolha *Deploy from a branch*, branch `main`, pasta `/ (root)`.
4. Na planilha: menu **Violão > Definir endereço do app (GitHub Pages)** e cole
   `https://netotec.github.io/acompanhamento-violao/`.

Daqui para frente, **Enviar link** gera links nesse endereço. Os links antigos continuam funcionando.

## Observações

- O endereço do web app dentro de `index.html` é o da implantação oficial. Se você criar outra
  implantação, troque a constante `EXEC` nos dois arquivos.
- O código do aparelho do aluno passa a ficar guardado na página do GitHub Pages. Quem já tinha
  aparelho cadastrado pelo link antigo conta como aparelho novo (limite de 2). Se o aluno ficar
  bloqueado, use **Alunos > Editar > Liberar aparelho novo**.
