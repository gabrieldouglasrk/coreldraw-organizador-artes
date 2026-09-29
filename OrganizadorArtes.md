# **Organizador de Artes para CorelDRAW**

### Automatize o fechamento de arquivos com muitas páginas no CorelDRAW

Macro VBA criado para quem recebe **PDFs com muitas páginas**, como 100, 200, 500 ou mais páginas, e precisa transformar essas páginas em uma montagem organizada para produção.

A solução é especialmente útil quando **todas as páginas possuem artes do mesmo tamanho**.

Em vez de organizar cada arte manualmente no CorelDRAW, o macro automatiza o processo e economiza tempo na preparação e no fechamento do arquivo.

## **Para que serve?**

Imagine receber um PDF com:

* 100 páginas
* 200 páginas
* 500 páginas
* ou uma quantidade ainda maior de páginas

E precisar colocar todas as artes em uma única mídia para produção.

Se todas as artes possuem o mesmo tamanho, fazer isso manualmente pode consumir bastante tempo.

O macro foi criado justamente para esse cenário.

### **PDF com muitas páginas → CorelDRAW → montagem automática**

O macro pega as artes que estão distribuídas nas páginas do CorelDRAW, redimensiona e organiza automaticamente em linhas e colunas.

A altura da mídia é calculada automaticamente de acordo com a quantidade de artes.

## **O que o macro faz**

* Organiza várias artes automaticamente.
* Trabalha com documentos com grande quantidade de páginas.
* Redimensiona as artes para o tamanho definido.
* Organiza as artes em linhas e colunas.
* Calcula automaticamente a altura necessária da mídia.
* Permite trabalhar sem espaço entre as artes.
* Permite definir margem.
* Adiciona uma borda mínima nas artes.
* Mantém as páginas originais do documento.
* Evita o uso de `Copy/Paste` para duplicar as artes.

## **Configuração atual**

* Mídia: 100 cm de largura
* Artes: 15,5 × 15,5 cm
* Espaçamento: 0 cm
* Margem: 0 cm
* Borda mínima
* Altura automática

### **Como usar**

1. Importe o PDF para o CorelDRAW.
2. Deixe cada arte em uma página.
3. Abra o VBA.
4. Importe ou cole `OrganizadorArtes.bas`.
5. Se necessário, altere os parâmetros no início do código.
6. Execute `OrganizarArtes_Configuravel`.

### **Configuração**

No início do código:

```vb
Const MIDIA_LARGURA_CM As Double = 100
Const ARTE_LARGURA_CM As Double = 15.5
Const ARTE_ALTURA_CM As Double = 15.5
Const ESPACO_CM As Double = 0
Const MARGEM_CM As Double = 0
Const USAR_BORDA As Boolean = True
Const TAMANHO_EXATO As Boolean = True
```

Basta alterar esses valores para outro serviço.

## **Exemplo**

Um PDF com **100 páginas**, contendo 100 artes do mesmo tamanho.

Em vez de organizar as 100 artes manualmente, o macro cria automaticamente uma nova página de montagem e distribui as artes conforme a largura da mídia.

```text
PDF
100 páginas
    ↓
CorelDRAW
    ↓
Organizador de Artes
    ↓
Montagem automática
    ↓
Arquivo pronto para produção
```

## **Objetivo**

O objetivo é **ganhar tempo na produção e no fechamento de arquivos**, principalmente em trabalhos com grande quantidade de artes iguais ou do mesmo tamanho.

A ferramenta foi criada a partir de uma necessidade real de produção gráfica e pode ser adaptada para diferentes tamanhos de mídia e de arte.

## **Palavras-chave**

CorelDRAW, VBA, macro CorelDRAW, organizar artes, fechamento de arquivo, produção gráfica, PDF, PDF com muitas páginas, PDF 100 páginas, PDF 200 páginas, PDF 500 páginas, organizar PDF no CorelDRAW, montagem de artes, fechamento de PDF, automação CorelDRAW, artes para impressão, preparação de arquivos, pré-impressão.
