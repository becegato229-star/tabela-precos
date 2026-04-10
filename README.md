# MUBEC — Tabela de Preços

Tabela de preços interativa gerada automaticamente a partir dos arquivos Excel.

## Como funciona

Toda vez que você atualizar um arquivo Excel na pasta `dados/` e fazer push para o GitHub, a tabela é regenerada e publicada automaticamente em ~1 minuto.

## Estrutura

```
dados/
  precos_suportes.xlsx   ← Preços dos suportes (NAT, GE, GF, INOX)
  precos_outros.xlsx     ← Porcas, arruelas, abraçadeiras, etc.
  embalagens.xlsx        ← Quantidades de embalagem por item
  logo.jpg               ← Logo da MUBEC
build.py                 ← Script que gera o HTML
```

## Configuração inicial (só uma vez)

1. Crie um repositório no GitHub (pode ser privado)
2. Faça upload de todos esses arquivos mantendo a estrutura de pastas
3. Vá em **Settings → Pages** e selecione:
   - Source: **Deploy from a branch**
   - Branch: **gh-pages** / **(root)**
4. Vá em **Settings → Actions → General** e garanta que Actions têm permissão de leitura e escrita
5. Faça qualquer push — a action vai rodar e publicar o site

## Atualizando preços

1. Abra o arquivo Excel correspondente
2. Edite os preços
3. Salve e faça upload no GitHub substituindo o arquivo antigo
4. Aguarde ~1 minuto — a tabela é atualizada automaticamente

## URL do site

Após a configuração, a tabela ficará disponível em:
`https://SEU_USUARIO.github.io/NOME_DO_REPOSITORIO`
