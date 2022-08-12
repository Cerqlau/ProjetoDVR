# ProjetoDVR

O código no VBA Excel que realiza o renomeio de qualquer arquivo por ordem de emissão, em uma pasta a escolha do usuário. Ela realiza o renomeio de acordo com um de padrão de renomeio, tendo como premissa de que há necessidade de organização dos produto final em ordem crescente, cabe ainda ressaltar que para este código os arquivos são a serem renomeados são gerados em intevalos definidos.

O intuito desta divulgação é que o código faz manipulação do "Scripting.FileSystemObject" o que pode ser facilmente adptado para outras atomatizações além de servir como objeto de estudo

## 🚀 Começando

Essas instruções permitirão que você obtenha uma cópia do projeto em operação na sua máquina local para fins de desenvolvimento e teste.

### 📋 Pré-requisitos

```
=> Excel 2019 ou superior;
=> Para melhor compatibilidade a execução das notificações, utilizar a marco no Windows 10 ou superior. 
```

### 🔧 Pré-configurações

1- O aplicativo do excel necessita estar com a opção "Habilitar Macros VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)

2- O aplicativo do excel necessita estar com a opção "Confiar no acesso ao modelo de projeto do VBA" habilitada ( Caminho: Opções> Central de confiabilidade > Configuração de Macro > Habilitar Macros VBA)

##  Premissas 

1- Todos os arquivos na pasta de destino serão renomeados independente da extensão. 

2- Todas as extensões de arquivos serão mantidas após renomeio. 

3- O renomeio dos arquivos é realizado de forma crescente de acordo a data de criação dos arquivos.

4- Os arquivos serão renomados a cada 30 min a partir da data e hora inicial 

   Nota: Tomar como base o primeiro arquivo para identificação do horário inicial

5-Os arquivos a serem renomeados não devem posuir nome idêntico ao renomeio que macro irá gerar. Aconselha-se modificar estes para um nome aleatório primeiramente.

### ⚙️ Executando o programa

1- Selecionar os vídeos a serem renomeados ( Nota: A duração de cada vídeo deverá ser de aprox. 30 min)

2- Copiar todos os vídeos em sequência de 30 min tendo como base a ultima data de modificação e separar em pastas distintas. 

3- Utilize a sub 1 para efetuar o renomeio dos arquivos. 

4- Utilize a sub 2 para replicar o renomeio dos arquivos para demais arquivos que necessitar ( respeitando quantidade existente e mesmo período de criação)

### 📨 Distribuição

É possivel efetuar a distribuição salvando os módulos em um pasta de trabalho habilitada para macros do vba. 

## 📦 Desenvolvimento

Lauro Cerqueira

LinkdIn: https://www.linkedin.com/in/lauro-cerqueira-70473568/

Instagram : laurorcerqueira

## 🛠️ Construído com

* [Microssoft Office Excel](https://docs.microsoft.com/pt-br/office/client-developer/excel/excel-home)
* [Visual Basic for Applications](https://docs.microsoft.com/pt-br/office/vba/api/overview/)

## 📄 Licença

Este projeto está sob a licença MIT - veja o arquivo [LICENSE.md](https://github.com/usuario/projeto/licenca) para detalhes.

## 🎁 

* Conte a outras pessoas sobre este projeto 📢
* Convide alguém da equipe para uma cerveja 🍺 
* Obrigado publicamente 🤓.

