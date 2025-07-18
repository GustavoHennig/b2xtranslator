
Tratar caracteres 0x13, 0x14 e 0x15 como no manual do MS-DOC.


Esse padrão aparece ao debugar o método: `ProcessHyperlinkFieldsInFallbackText`

Eu estava parseando um fragmento de um arquivo doc, na parte dos links, e vi essa sequência, isso é um padrão esperado do doc, certo? Me refiro ao \u0013, \u0014 e \u0015 (parecem separadores)
Note que há um erro no doc original, Um link que aparentemente é um são 2, um grudado ao outro.

```
"Apache Tika: \u0013HYPERLINK \"http://tika.apache.org/\" \\h\u0014http://tika.apache.org/\u0015 \u0013HYPERLINK \"http://tika.apache.org/\" \\h\u0014Tika\u0015\r"
```

Deveria ser interpretado da seguinte forma:

Isso significa:

Primeiro campo HYPERLINK:

\u0013 → Início do campo

HYPERLINK "http://tika.apache.org/" \h → Instrução do campo (target e parâmetro \h que indica link do tipo hyperlink "hot")

\u0014 → Separador (a partir daqui é o que é mostrado ao usuário se o campo estiver expandido)

http://tika.apache.org/ → texto exibido

\u0015 → Fim do campo

Segundo campo HYPERLINK colado:

\u0013HYPERLINK "http://tika.apache.org/" \h

\u0014Tika

\u0015



## Scenario
`Bug51686.doc`

The expected result is `Display Name (URL)`