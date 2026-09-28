# 38º Seminário IPBM — atualização do site e inscrição

## Site
Baseado na versão atual do site do 38º Seminário IPBM.

### Novos dados no formulário
- Instituição / Força
- UF da corporação para PM/CBM de outros Estados
- Posto / Graduação correspondente à força
- Abreviação é enviada pelo formulário e usada no certificado

### Regras do certificado
O certificado passa a usar uma única identificação na caixa `{{NOME}}` do modelo Google Slides:

`POSTO/GRADUAÇÃO + NOME`

Ex.: `3º Sgt PM EVERTON ALVES FERNANDES`

## Apps Script
Use o `apps-script/Codigo.gs` desta entrega.
Mantenha a sua chave administrativa atual em `ADMIN_KEY`.

O modelo Google Slides continua sendo a apresentação:
https://docs.google.com/presentation/d/1F9qzntKj9qa5ZBFVWIDtbQp3PAwbQOL-oq4WexfkYLU/edit

Antes do teste, confirme que a apresentação-modelo contém o marcador `{{NOME}}` na caixa de texto do nome.

## Implantação
1. Atualize o `index.html`, `script.js` e `admin.html` no GitHub Pages.
2. No Apps Script, substitua somente `Código.gs` pelo arquivo desta entrega.
3. Mantenha `admin.html` e `certificate.html` atuais.
4. Salve e publique uma nova versão do Web App na mesma implantação.
5. Para o certificado de teste, bloqueie e libere novamente a autorização para limpar o PDF antigo.

## Observação sobre inscrições antigas
Os registros existentes continuam válidos. Os novos campos `UF da corporação` e `Posto / Graduação` ficarão vazios nas inscrições antigas até serem preenchidos.
