# lanca-nota-uff
Código para lançamento automatizado de notas no sistema IdUFF

## São necessárias as seguintes bibliotecas do Python:
* selenium
* openpyxl
* pandas

## Passos:
* Certifique-se de que tem o Google Chrome instalado e verifique a versão
* Baixe o ChromeDriver em https://chromedriver.chromium.org/, atentando à versão do seu Google Chrome, e salve o arquivo `chromedrive.exe` na mesma pasta do código Python
* Prepare o arquivo (.xlsx) com o número de matrícula, numa coluna com título Matricula (sem acento), e nota de cada aluno, numa coluna com título Nota (ver exemplo Notas.xlsx)
* Na seção de `# === INPUTS === #` do código, especifique o nome do arquivo onde serão salvas as notas retiradas do site, se houver. Ele pode ser utilizado, posteriormente, para checagem das notas lançadas
* Na seção de `# === INPUTS === #` do código, especifique o nome do arquivo com as notas
* Execute o código e aguarde a janela do Chrome abrir
* Entre com seu login e senha
* Clique na turma onde as notas serão lançadas
* Pressione Enter no prompt de onde está sendo executado o código Python
* Aguarde até que seja exibida a lista de alunos e notas (se houver) do sistema
* Digite `S` (sim) para prosseguir (ou `N` para cancelar)
* Aguarde a exibição das notas que serão lançadas, conforme matrículas do site que foram encontradas no arquivo .xlsx
* Confira as notas e, caso estejam corretas, corretas, digite `S` e pressione enter
* Aguarde o lançamento das notas e verifique os valores na janela do Chrome

## Atenção!
* Este código tem como objetivo agilizar o processo de lançamento de notas, mas a responsabilidade dos valores registrados é do usuário
