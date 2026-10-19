# Painel de Alocação de Guichês

Painel estático para organizar funcionários em cinco guichês, nos períodos da manhã e da tarde.

## Executar localmente

No uso diário, dê dois cliques em `Abrir Painel.bat`. O servidor será iniciado minimizado e o navegador abrirá automaticamente.

Como alternativa, abra um terminal nesta pasta e execute:

```powershell
& 'C:\Users\IvisonSena\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' .\server.mjs
```

Mantenha o terminal aberto e acesse `http://127.0.0.1:8080/`.

## Alimentação pelo Google Sheets

As abas `Alocações`, `Cursos`, `Folgas` e `Contatos` já estão conectadas aos respectivos endereços publicados da planilha. O administrador edita os dados no Google Planilhas e os visitantes usam o botão **Atualizar** para buscar as mudanças.

## Publicar no GitHub Pages

1. Envie os arquivos desta pasta para a raiz de um repositório GitHub.
2. No repositório, abra **Settings → Pages**.
3. Em **Build and deployment**, escolha **Deploy from a branch**.
4. Selecione a branch principal, a pasta **/(root)** e salve.

O arquivo `server.mjs` não é utilizado pelo GitHub Pages. O painel publicado funciona como site estático e consulta diretamente os CSVs públicos do Google Planilhas.
