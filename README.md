# **Organizador de Artes para CorelDRAW**

Macro VBA para CorelDRAW que automatiza a organização de **grandes quantidades de artes do mesmo tamanho** em uma única mídia.

## **O problema**

Quando um PDF possui 100, 200 ou mais páginas com artes do mesmo tamanho, é necessário repetir manualmente tarefas como copiar, redimensionar, posicionar e, em alguns casos, utilizar PowerClip.

Com muitas páginas, esse processo se torna repetitivo e burocrático.

## **A solução**

O macro automatiza esse processo diretamente no CorelDRAW, aproveitando os recursos do **VBA**, sem depender de ferramentas externas.

Ele:

* Redimensiona as artes.
* Organiza em linhas e colunas.
* Calcula automaticamente a altura da mídia.
* Permite definir espaçamento e margem.
* Adiciona borda mínima.
* Mantém as páginas originais.
* Não utiliza `Copy/Paste` para duplicar as artes.

**PDF com muitas páginas → CorelDRAW → montagem automática → produção.**

## **Configuração atual**

* Mídia: **100 cm de largura**
* Artes: **15,5 × 15,5 cm**
* Espaçamento: **0 cm**
* Margem: **0 cm**
* Borda: **mínima**
* Altura: **automática**

### **Como usar**

1. Importe o PDF para o CorelDRAW, deixando uma arte por página.
2. Abra o VBA.
3. Importe ou cole `OrganizadorArtes.bas`.
4. Se necessário, altere os parâmetros no início do código.
5. Execute `OrganizarArtes_Configuravel`.

### **Configuração**

```vb id="8j1s8j"
Const MIDIA_LARGURA_CM As Double = 100
Const ARTE_LARGURA_CM As Double = 15.5
Const ARTE_ALTURA_CM As Double = 15.5
Const ESPACO_CM As Double = 0
Const MARGEM_CM As Double = 0
Const USAR_BORDA As Boolean = True
Const TAMANHO_EXATO As Boolean = True
```

Basta alterar esses valores para outro serviço.

## **Objetivo**

Ganhar tempo na **produção e no fechamento de arquivos**, principalmente quando existe uma grande quantidade de artes iguais ou do mesmo tamanho.

O projeto nasceu de uma necessidade real de produção gráfica e busca aproveitar o VBA disponível no próprio CorelDRAW para automatizar tarefas repetitivas.

## **Tecnologias**

* CorelDRAW
* VBA (Visual Basic for Applications)

**Palavras-chave:** CorelDRAW, VBA, macro CorelDRAW, PDF, PDF 100 páginas, PDF 200 páginas, organizar PDF, fechamento de arquivo, produção gráfica, montagem de artes, automação CorelDRAW, pré-impressão.
