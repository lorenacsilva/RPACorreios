# RPA - Correios

RPA (Robotic Process Automation) criado utilizando Selinium como framework e ClosedXML como .Net Libray, o projeto foi desenvolvido como automação com seguintes passos:
1. Acessar site dos correios https://buscacepinter.correios.com.br/app/endereco/index.php
2. Fazer busca de CEP
3. Trazer dados da tabela para uma planilha do excel

### Configurações de projeto

Para utilizar **Selinium** é necessário instalação de alguns pacotes específicos no nuget, são eles **Selenium.WebDriver** e **Selenium.WebDriver.ChromeDriver**, o segundo pacote irá depender do navegador que irá utilizar. 
<div> No meu caso utilizei Chrome Version 85 porque essa biblioteca suporta apenas essa versão. A seguir alguns links me ajudaram durante a reinstalação do navegador:</div>

<div>

- Como baixar: https://browserhow.com/how-to-downgrade-and-install-older-version-of-chrome/
- Onde baixar: https://www.filepuma.com/download/google_chrome_64bit_85.0.4183.83-26533/
- Como desativar atualizações automáticas:https://www.apptuts.net/tutorial/web/desativar-atualizacoes-automaticas-chrome-pc/
  
  </div>
  
 No caso do **CloseXML** é bem mais tranquilo por ser uma biblioteca onde todas as configurações, ajudas e demais detalhes estão dispoíveis no próprio github.
<div> A seguir alguns links e comandos que me ajudaram durante o desenvolvimento</div>

<div>

-  Instalar pacote do XML utilizando Nuget ou linha de comando ``Install-Package closedxml `` 
-  Bibliotca: https://github.com/ClosedXML/ClosedXML
-  Como utilizar formulas: https://github.com/ClosedXML/ClosedXML/wiki/Evaluating-Formulas
  
  </div>
