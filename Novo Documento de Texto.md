Projeto: Organizador de Artes para Plotagem no CorelDRAW
1. Objetivo do projeto

Criamos um macro em VBA para CorelDRAW cuja função é pegar artes que estão distribuídas, normalmente uma arte por página, e criar uma nova página de montagem contendo todas as artes organizadas automaticamente.

A ideia é transformar isso em um macro reutilizável para diferentes serviços, permitindo alterar facilmente apenas alguns parâmetros no início do código.
2. Funcionamento definido

O macro deve:

    Ler todas as páginas originais do documento.

    Ignorar páginas que estejam vazias.

    Considerar cada página com conteúdo como uma arte.

    Duplicar a arte sem usar Copy/Paste.

    Colocar as cópias em uma nova página de montagem.

    Redimensionar as artes conforme os parâmetros definidos.

    Organizar automaticamente as artes em linhas e colunas.

    Usar toda a largura disponível da mídia.

    Calcular automaticamente a altura necessária.

    Não limitar a altura a 100 cm.

    Criar uma página tão alta quanto for necessário para acomodar todas as artes.

    Adicionar uma borda extremamente fina em cada arte.

    Não deixar espaço entre as artes.

    Não deixar margem externa.

    Manter as páginas originais no documento.

3. Parâmetros atuais do serviço
Mídia

Largura: 100 cm

Altura: automática.

A altura não é mais definida manualmente. O VBA calcula:

    quantidade de linhas × altura da arte

Portanto, se houver muitas artes, a página simplesmente fica mais comprida.
Artes

Cada arte:

15,5 × 15,5 cm

Como você informou que as artes utilizadas nesse serviço são quadradas, podemos usar tamanho exato.
Organização

Espaçamento:

0 cm

Ou seja, as artes ficam encostadas.

Margem:

0 cm

Ou seja, a primeira arte começa na própria borda da mídia.
Quantidade por linha

Com mídia de 100 cm e artes de 15,5 cm:

    6 artes = 93 cm

    7 artes = 108,5 cm → não cabe

Portanto:

6 artes por linha.

Exemplo:

┌──────┬──────┬──────┬──────┬──────┬──────┐
│ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │
├──────┼──────┼──────┼──────┼──────┼──────┤
│ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │
├──────┼──────┼──────┼──────┼──────┼──────┤
│ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │ 15,5 │
└──────┴──────┴──────┴──────┴──────┴──────┘
              100 cm

4. Borda

Você pediu uma borda em cada arte.

Decidimos deixar a borda fixa e extremamente fina, sem necessidade de editar esse parâmetro toda vez.

Valor utilizado:

0,05 mm

A borda:

    não possui preenchimento;

    possui apenas contorno;

    é criada exatamente no tamanho da arte;

    fica atrás da arte.

5. Problema que encontramos com Copy/Paste

A primeira versão utilizava:

sr.Copy

e depois:

Paste

O Corel apresentou:

Erro em tempo de execução '-2147467259 (80004005)'

A depuração apontou especificamente para:

sr.Copy

Por isso mudamos a arquitetura.

Agora o macro usa:

Set srCopia = sr.Duplicate

e:

srCopia.MoveToLayer pgDestino.ActiveLayer

Assim não depende da área de transferência do Corel/Windows.
6. Versatilidade do projeto

A ideia final é que você não precise alterar o código inteiro quando receber outro serviço.

No início do macro ficam os parâmetros:

Const MIDIA_LARGURA_CM As Double = 100

Const ARTE_LARGURA_CM As Double = 15.5
Const ARTE_ALTURA_CM As Double = 15.5

Const ESPACO_CM As Double = 0
Const MARGEM_CM As Double = 0

Const USAR_BORDA As Boolean = True

Const TAMANHO_EXATO As Boolean = True

Por exemplo, futuramente para uma arte de 20 × 10 cm em uma mídia de 120 cm:

Const MIDIA_LARGURA_CM As Double = 120

Const ARTE_LARGURA_CM As Double = 20
Const ARTE_ALTURA_CM As Double = 10

O restante do cálculo é automático.
7. Código consolidado

Importante: ao copiar para o VBA, copie somente o conteúdo que começa em Sub e termina em End Sub. Não copie os ``` nem qualquer identificação que apareça fora do código.

Sub OrganizarArtes_Configuravel()

    '=========================================================
    ' CONFIGURACOES DO SERVICO
    ' ALTERE SOMENTE OS PARAMETROS ABAIXO
    '=========================================================

    ' LARGURA DA MIDIA EM CM
    Const MIDIA_LARGURA_CM As Double = 100

    ' TAMANHO DA ARTE EM CM
    Const ARTE_LARGURA_CM As Double = 15.5
    Const ARTE_ALTURA_CM As Double = 15.5

    ' ESPACO ENTRE AS ARTES EM CM
    ' 0 = artes encostadas
    Const ESPACO_CM As Double = 0

    ' MARGEM DA MIDIA EM CM
    ' 0 = sem margem
    Const MARGEM_CM As Double = 0

    ' CRIAR BORDA?
    ' True = sim
    ' False = nao
    Const USAR_BORDA As Boolean = True

    ' TAMANHO EXATO?
    ' True = usa exatamente largura x altura informadas
    ' False = mantem a proporcao original
    Const TAMANHO_EXATO As Boolean = True


    '=========================================================
    ' FIM DAS CONFIGURACOES
    ' NAO PRECISA ALTERAR ABAIXO
    '=========================================================

    Dim doc As Document
    Dim pgOrigem As Page
    Dim pgDestino As Page

    Dim sr As ShapeRange
    Dim srCopia As ShapeRange

    Dim arte As Shape
    Dim borda As Shape

    Dim qtdPaginasOriginais As Long
    Dim qtdArtes As Long
    Dim i As Long

    Dim larguraPagina As Double
    Dim alturaPagina As Double

    Dim larguraDesejada As Double
    Dim alturaDesejada As Double

    Dim margem As Double
    Dim espaco As Double

    Dim x As Double
    Dim y As Double

    Dim coluna As Long
    Dim linha As Long

    Dim artesPorLinha As Long
    Dim totalLinhas As Long

    Dim escala As Double

    Set doc = ActiveDocument


    '=========================================================
    ' CONVERTER CM PARA A UNIDADE DO CORELDRAW
    '=========================================================

    larguraPagina = Application.ConvertUnits( _
        MIDIA_LARGURA_CM * 10, _
        cdrMillimeter, _
        doc.Unit)

    larguraDesejada = Application.ConvertUnits( _
        ARTE_LARGURA_CM * 10, _
        cdrMillimeter, _
        doc.Unit)

    alturaDesejada = Application.ConvertUnits( _
        ARTE_ALTURA_CM * 10, _
        cdrMillimeter, _
        doc.Unit)

    margem = Application.ConvertUnits( _
        MARGEM_CM * 10, _
        cdrMillimeter, _
        doc.Unit)

    espaco = Application.ConvertUnits( _
        ESPACO_CM * 10, _
        cdrMillimeter, _
        doc.Unit)


    '=========================================================
    ' CONTAR PAGINAS ORIGINAIS
    '=========================================================

    qtdPaginasOriginais = doc.Pages.Count

    If qtdPaginasOriginais = 0 Then

        MsgBox "Nao existem paginas no documento.", vbExclamation

        Exit Sub

    End If


    '=========================================================
    ' CONTAR SOMENTE PAGINAS COM CONTEUDO
    '=========================================================

    qtdArtes = 0

    For i = 1 To qtdPaginasOriginais

        If doc.Pages(i).Shapes.Count > 0 Then
            qtdArtes = qtdArtes + 1
        End If

    Next i


    If qtdArtes = 0 Then

        MsgBox "Nenhuma arte foi encontrada.", vbExclamation

        Exit Sub

    End If


    '=========================================================
    ' CALCULAR QUANTAS ARTES CABEM POR LINHA
    '=========================================================

    artesPorLinha = Int( _
        (larguraPagina - (margem * 2) + espaco) / _
        (larguraDesejada + espaco))


    If artesPorLinha < 1 Then

        MsgBox "A arte e maior que a largura da midia.", vbCritical

        Exit Sub

    End If


    '=========================================================
    ' CALCULAR QUANTIDADE DE LINHAS
    '=========================================================

    totalLinhas = _
        Int((qtdArtes + artesPorLinha - 1) / artesPorLinha)


    '=========================================================
    ' CALCULAR ALTURA TOTAL DA MIDIA
    '
    ' A ALTURA E AUTOMATICA.
    ' NAO EXISTE LIMITE DE 100 CM.
    '=========================================================

    alturaPagina = _
        (margem * 2) + _
        (totalLinhas * alturaDesejada) + _
        ((totalLinhas - 1) * espaco)


    '=========================================================
    ' CRIAR PAGINA DE MONTAGEM
    '
    ' LARGURA = 100 CM
    ' ALTURA = CALCULADA AUTOMATICAMENTE
    '=========================================================

    Set pgDestino = doc.AddPages(1)

    pgDestino.SizeWidth = larguraPagina
    pgDestino.SizeHeight = alturaPagina

    pgDestino.Activate


    '=========================================================
    ' PROCESSAR CADA PAGINA ORIGINAL
    '=========================================================

    qtdArtes = 0

    For i = 1 To qtdPaginasOriginais

        Set pgOrigem = doc.Pages(i)


        '-----------------------------------------------------
        ' IGNORAR PAGINA VAZIA
        '-----------------------------------------------------

        If pgOrigem.Shapes.Count > 0 Then

            qtdArtes = qtdArtes + 1


            '-------------------------------------------------
            ' PEGAR OBJETOS DA PAGINA ORIGINAL
            '-------------------------------------------------

            Set sr = pgOrigem.Shapes.All


            '-------------------------------------------------
            ' DUPLICAR OBJETOS
            '
            ' NAO USA COPY/PASTE
            '-------------------------------------------------

            Set srCopia = sr.Duplicate


            '-------------------------------------------------
            ' MOVER COPIA PARA A PAGINA DE MONTAGEM
            '-------------------------------------------------

            srCopia.MoveToLayer pgDestino.ActiveLayer


            '-------------------------------------------------
            ' AGRUPAR A ARTE
            '-------------------------------------------------

            Set arte = srCopia.Group


            '=================================================
            ' AJUSTAR TAMANHO
            '=================================================

            If TAMANHO_EXATO = True Then

                arte.SetSize _
                    larguraDesejada, _
                    alturaDesejada

            Else

                If arte.SizeWidth > 0 Then

                    escala = _
                        larguraDesejada / arte.SizeWidth

                    arte.SetSize _
                        larguraDesejada, _
                        arte.SizeHeight * escala

                End If

            End If


            '=================================================
            ' CALCULAR POSICAO
            '=================================================

            coluna = _
                (qtdArtes - 1) Mod artesPorLinha

            linha = _
                (qtdArtes - 1) \ artesPorLinha


            x = margem + _
                coluna * (larguraDesejada + espaco)


            y = alturaPagina - margem - _
                linha * (alturaDesejada + espaco)


            '=================================================
            ' POSICIONAR ARTE
            '=================================================

            arte.LeftX = x
            arte.TopY = y


            '=================================================
            ' CRIAR BORDA
            '=================================================

            If USAR_BORDA = True Then

                Set borda = _
                    pgDestino.ActiveLayer.CreateRectangle( _
                        arte.LeftX, _
                        arte.TopY, _
                        arte.RightX, _
                        arte.BottomY)


                ' Sem preenchimento
                borda.Fill.ApplyNoFill


                ' Borda extremamente fina
                ' 0,05 mm

                borda.Outline.Width = _
                    Application.ConvertUnits( _
                        0.05, _
                        cdrMillimeter, _
                        doc.Unit)


                ' Colocar a borda atras da arte
                borda.OrderToBack

            End If

        End If

    Next i


    '=========================================================
    ' ATIVAR PAGINA FINAL
    '=========================================================

    pgDestino.Activate


    '=========================================================
    ' MENSAGEM FINAL
    '=========================================================

    MsgBox _
        "Concluido!" & vbCrLf & vbCrLf & _
        "Artes encontradas: " & qtdArtes & vbCrLf & _
        "Midia: " & MIDIA_LARGURA_CM & " cm de largura" & vbCrLf & _
        "Altura calculada automaticamente" & vbCrLf & _
        "Arte: " & ARTE_LARGURA_CM & " x " & _
            ARTE_ALTURA_CM & " cm" & vbCrLf & _
        "Artes por linha: " & artesPorLinha & vbCrLf & _
        "Linhas: " & totalLinhas & vbCrLf & _
        "Espacamento: " & ESPACO_CM & " cm" & vbCrLf & _
        "Margem: " & MARGEM_CM & " cm" & vbCrLf & _
        "Borda: minima", _
        vbInformation, _
        "Organizacao concluida"


End Sub

8. Como usar no futuro

Você abre o VBA e procura apenas esta parte:

Const MIDIA_LARGURA_CM As Double = 100

Const ARTE_LARGURA_CM As Double = 15.5
Const ARTE_ALTURA_CM As Double = 15.5

Const ESPACO_CM As Double = 0
Const MARGEM_CM As Double = 0

Const USAR_BORDA As Boolean = True

Const TAMANHO_EXATO As Boolean = True

Exemplo: outro trabalho

Se forem artes de 10 × 20 cm em mídia de 80 cm:

Const MIDIA_LARGURA_CM As Double = 80

Const ARTE_LARGURA_CM As Double = 10
Const ARTE_ALTURA_CM As Double = 20

O macro calcula sozinho quantas cabem por linha e qual será a altura final.
Regra principal do projeto

A largura da mídia é fixa. A altura nunca é fixa: ela é calculada conforme a quantidade de artes.

E o problema que encontramos com sr.Copy foi resolvido definitivamente pela utilização de Duplicate + MoveToLayer, evitando a área de transferência.

Observação: esta é a versão consolidada do que construímos na conversa; como houve várias iterações e testes, eu trataria essa versão como a base do projeto, mas faria um teste com uma cópia do arquivo antes de usar em produção.