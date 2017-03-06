# <a name="office-add-in-that-converts-directly-between-word-and-markdown-formats"></a>Suplemento do Office que converte diretamente entre os formatos Word e Markdown

Use as APIs do Word.js para converter um documento Markdown para o Word para edição e converta o documento do Word para o formato Markdown usando os objetos Paragraph, Table, List e Range.

![Converter entre Word e Markdown](readme_art/ReadMeScreenshot.PNG)

## <a name="table-of-contents"></a>Sumário
* [Histórico de alterações](#change-history)
* [Pré-requisitos](#prerequisites)
* [Testar o suplemento](#test-the-add-in)
* [Problemas conhecidos](#known-issues)
* [Perguntas e comentários](#questions-and-comments)
* [Recursos adicionais](#additional-resources)

## <a name="change-history"></a>Histórico de alterações

16 de dezembro de 2016:

* Versão inicial.

## <a name="prerequisites"></a>Pré-requisitos

* Visual Studio 2015 ou posterior.
* Word 2016 para Windows, build 16.0.6727.1000 ou superior.

## <a name="test-the-add-in"></a>Testar o suplemento

1. Clone ou baixe o projeto para sua área de trabalho.
2. Abra o arquivo Word-Add-in-JavaScript-MDConversion.sln no Visual Studio.
2. Pressione F5.
3. Depois que o Word for iniciado, pressione o botão **Abrir Conversor** na faixa de opções **Página Inicial**.
4. Quando o aplicativo for carregado, pressione o botão **Inserir documento Markdown de teste**.
5. Depois que o texto Markdown de exemplo for carregado, pressione o botão **Converter texto MD para o Word**.
6. Depois que o documento tiver sido convertido para o Word, edite-o. 
7. Pressione o botão **Converter documento para Markdown**. 
8. Depois que o documento for convertido, copie e cole seu conteúdo em um visualizador para Markdown, como o Visual Studio Code.
9. Como alternativa, você pode começar com o botão **Inserir documento Word de teste** e converter o documento Word de exemplo criado para Markdown. 
10. Opcionalmente, comece com seu próprio texto Markdown ou conteúdo do Word e teste o suplemento.

## <a name="known-issues"></a>Problemas conhecidos

- Devido a um bug na maneira que as listas do Word criadas por programação são criadas, o Markdown para Word só converterá corretamente a primeira lista (ou às vezes as duas primeiras listas) em um documento. (Qualquer número de listas do Markdown será convertido corretamente para o Word.)
- Quando você converte o mesmo documento repetidamente entre o Word e o Markdown, todas as linhas das tabelas assumem a formatação da linha de cabeçalho, que geralmente inclui texto em negrito.
- O suplemento usa algumas APIs do Office que ainda não têm suporte no Word Online (desde 15/02/2017). Você deve testá-lo no Word para área de trabalho (pressione F5 para abri-lo automaticamente).

## <a name="questions-and-comments"></a>Perguntas e comentários

Gostaríamos de saber sua opinião sobre este exemplo. Você pode nos enviar comentários na seção *Issues* deste repositório.

As perguntas sobre o desenvolvimento do Microsoft Office 365 em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Se sua pergunta estiver relacionada às APIs JavaScript para Office, não deixe de marcá-la com as tags [office-js] e [API].

## <a name="additional-resources"></a>Recursos adicionais

* [Documentação dos suplementos do Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro de Desenvolvimento do Office](http://dev.office.com/)
* Confira outros exemplos de Suplemento do Office em [OfficeDev no Github](https://github.com/officedev)

## <a name="copyright"></a>Direitos autorais
Copyright (C) 2016 Microsoft Corporation. Todos os direitos reservados.

