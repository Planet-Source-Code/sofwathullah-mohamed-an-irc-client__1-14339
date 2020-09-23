Attribute VB_Name = "Module1"
' This forces you to define all variables, good
' programming practice as my teacher told me
' Option Explicit
' This simply says any variable not defined
' a type, will default to integer
DefInt A-Z

'Global variables


Sub xlate(a$)
'MsgBox a$
' ctype=0 means akuru;  ctype=1 means onlu fili;    ctype=2 means akuru with fili
ak = 0: f = 1: af = 2
'text$ = ""
lin$ = ""
ln = Len(a$)
a$ = " " + LCase$(a$) + " "

'a$ = " " + LCase$(RTrim$(a$)) + " "

For lft = 2 To Len(a$) - 1
        lett0$ = Mid$(a$, lft - 1, 1)
        lett1$ = Mid$(a$, lft, 1)
        lett2$ = Mid$(a$, lft + 1, 1)
        lett3$ = Mid$(a$, lft + 2, 1)
    
    akuru$ = lett1$
    
    
    If Mid$(a$, lft - 1, 1) = " " And Mid$(a$, lft, 1) = " " Or Mid$(a$, lft + 1, 1) = " " And Mid$(a$, lft, 1) = " " Then GoTo here
    If lett1$ = "d" Then akuru$ = "D": ' GoTo here                      'daviyani
    
    
    If Mid$(a$, lft, 1) = "'" And Mid$(a$, lft + 1, 1) = "h" Then akuru$ = "wq": lft = lft + 1: ctype = af: GoTo here   'alifusukun
    
    If Asc(Mid$(a$, lft, 1)) <= 57 And Asc(Mid$(a$, lft, 1)) >= 33 Then akuru$ = lett1$: GoTo here
        'If form1.Image5.Tag = "print" Then akuru$ = lett1$ + " " Else akuru$ = lett1$
        'akuru$ = Mid$(a$, lft, 1):  GoTo here
        'MsgBox form1.Image5.Tag
        'lft = lft + 1
        'GoTo here
    'End If
    
    
    
    If Mid$(a$, lft, 1) = "e" And Mid$(a$, lft + 1, 1) = "y" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wE": lft = lft + 1: ctype = af: GoTo here ' alif + eybeyfili
    If Mid$(a$, lft, 1) = "i" And Mid$(a$, lft + 1, 1) = "y" And InStr("aeiouy", Mid$(a$, lft + 2, 1)) = 0 Then akuru$ = "tq": lft = lft + 1: ctype = af:  GoTo here 'thaa sukun
    If lett1$ = "a" And lett2$ = "a" And InStr("aeiou " + Chr$(34), lett0$) <> 0 Then akuru$ = "wA": lft = lft + 1: ctype = af: GoTo here        'alifu aabaafili
    If Mid$(a$, lft, 1) = "a" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wa": ctype = af: GoTo here 'lft = lft + 1    'alif abafili
    If Mid$(a$, lft, 1) = "u" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wu": ctype = af: GoTo here 'lft = lft + 1             'alifu ubufili
    '**********
    If InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 And Mid$(a$, lft, 1) = "o" And Mid$(a$, lft + 1, 1) = "a" Then akuru$ = "wO": lft = lft + 1: ctype = f:  GoTo here 'oaboafili
    
    If Mid$(a$, lft, 1) = "o" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wo": ctype = af: GoTo here                             'alifu obofili
    If lett1$ = "t" Then akuru$ = "T": ctype = ak ': lft = lft + 1 : GoTo here                                                           'taviyani
    If Mid$(a$, lft, 1) = "t" And Mid$(a$, lft + 1, 1) = "h" Then akuru$ = "t": lft = lft + 1: ctype = af                               'thaa
    If lett1$ = "t" And lett2$ = "h" And lett3$ = " " Then akuru$ = "tq":: GoTo here  'lft = lft + 1                        'thaasukun
        If Mid$(a$, lft, 1) = "k" And Mid$(a$, lft + 1, 1) = "h" Then akuru$ = "K":::                            ' GoTo here                             'khaa
        If lett1$ = "g" And lett2$ = "h" Then akuru$ = "G":: ctype = ak: GoTo here                                                      'ghainu
    If Mid$(a$, lft, 1) = "q" Then akuru$ = "Q"::               'GoTo here' lft = lft + 1                                              'qaafu
    If Mid$(a$, lft, 1) = "q" And InStr("aeiou", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = "Qq": ctype = af                               'qaafusukun
    'If Mid$(a$, lft, 1) = "t" And InStr("aeiou", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = "Tq": GoTo here: ctype = af                              'taviyanisukun
    If lett1$ = "d" And lett2$ = "h" And Mid$(a$, lft + 2, 1) = "d" And Mid$(a$, lft + 3, 1) = "h" Then akuru$ = "wq": lft = lft + 1: ctype = af: GoTo here 'alifusukun
    If lett1$ = "d" And lett2$ = "h" Then akuru$ = "d": lft = lft + 1: ctype = ak: GoTo here                                            'dhaalu
    If lett1$ = "a" And lett2$ = "a" Then akuru$ = "A": lft = lft + 1: ctype = f: GoTo here                            'aabaafili
    
    'If Mid$(a$, lft, 1) <> " " And Mid$(a$, lft + 1, 1) = "'" Then
    
    If InStr("aeiou", Mid$(a$, lft, 1)) = 0 And Mid$(a$, lft + 1, 1) = "'" Then
         'If Form1.Image5.Tag = "print" Then akuru$ = lett1$ + "  " Else
         akuru$ = lett1$
        'MsgBox form1.Image5.Tag
        lft = lft + 1: GoTo here
    End If
    
    If Mid$(a$, lft, 1) = "e" And Mid$(a$, lft + 1, 1) = "y" Then akuru$ = "E": lft = lft + 1: ctype = f: GoTo here    'eybeyfili
    If lett1$ = "l" And lett2$ = "h" And Mid$(a$, lft + 2, 1) = "l" And Mid$(a$, lft + 3, 1) = "h" Then akuru$ = "wq": lft = lft + 1: ctype = af: GoTo here 'alifusukun
    If Mid$(a$, lft, 1) = "l" And Mid$(a$, lft + 1, 1) = "h" Then akuru$ = "L": lft = lft + 1: ctype = ak: GoTo here    'lhaviyani
    If Mid$(a$, lft, 1) = "n" And InStr("aeiou", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = "nq": ctype = af: GoTo here    'noonusukun
    If Mid$(a$, lft, 1) = "i" And (InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1))) <> 0 Then akuru$ = "wi": ctype = af: GoTo here 'alifuibifili
    If Mid$(a$, lft, 1) = "e" And Mid$(a$, lft + 1, 1) = "h" And Mid$(a$, lft + 2, 1) = " " And InStr("aeiou ", Mid$(a$, lft - 1, 1)) = 0 Then akuru$ = akuru$ + "wq": lft = lft + 1: ctype = af: GoTo here  'alifusukun
    If Mid$(a$, lft, 1) = "e" And Mid$(a$, lft + 1, 1) = "e" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wI": lft = lft + 1: ctype = af: GoTo here  'alifu eebeefili
    If Mid$(a$, lft, 1) = "e" And Mid$(a$, lft + 1, 1) = "e" Then akuru$ = "I": lft = lft + 1: ctype = f: GoTo here    'eebeefili
    If Mid$(a$, lft, 1) = "e" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "we": ctype = af: GoTo here 'lft = lft + 1  'alifu ebefili
    If Mid$(a$, lft, 1) = "u" And InStr("aeiou " + Chr$(34), Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = "wu": ctype = af: GoTo here 'lft = lft + 1  'alifu ubufili
    If Mid$(a$, lft - 1, 1) = "a" And Mid$(a$, lft, 1) = "h" And InStr(" .,", Mid$(a$, lft + 1, 1)) <> 0 Then akuru$ = "Sq": lft = lft + 0: ctype = af: GoTo here   'shaviyani sukun
    
    If Mid$(a$, lft - 1, 1) = "o" And Mid$(a$, lft, 1) = "h" And InStr("aeiou ", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = "Sq": ctype = af:: lft = lft + 1: GoTo here     'shaviyani sukun
    '***************
    If Mid$(a$, lft, 1) = "o" And Mid$(a$, lft + 1, 1) = "a" Then akuru$ = "O": lft = lft + 1: ctype = f:  GoTo here 'oaboafili
    
    If lett1$ = "m" And lett2$ = " " Then akuru$ = "mq": ctype = af: GoTo here ' lft = lft + 1                         'meemusukun
    If lett1$ = "s" And lett2$ = "h" And Mid$(a$, lft + 2, 1) = "s" And Mid$(a$, lft + 3, 1) = "h" Then akuru$ = "wq": lft = lft + 1: ctype = af: GoTo here    'alifusukun
    
    If lett1$ = "c" And lett2$ = "h" And Mid$(a$, lft + 2, 1) = "c" And Mid$(a$, lft + 3, 1) = "h" Then akuru$ = "wq": lft = lft + 1: ctype = af: GoTo here    'alifusukun
    
    If lett1$ = "s" And lett2$ = "h" And InStr("aeiou", Mid$(a$, lft + 2, 1)) = 0 Then akuru$ = "Cq": lft = lft + 1: ctype = af: GoTo here                      'sheenu sukun
    If lett1$ = "s" And lett2$ = "h" Then akuru$ = "C": lft = lft + 1:: GoTo here                                      'sheenu
    If lett1$ = "z" And lett2$ = "h" Then akuru$ = "S": lft = lft + 1: ctype = ak: GoTo here                           'seenu
    If lett1$ = "s" And lett2$ = "h" And InStr("aeiou", lett3$) = 0 Then akuru$ = "Cq": ctype = af: GoTo here 'lft = lft + 1 'sheenu sukun
    If lett1$ = "s" And InStr("aeiou", lett2$) = 0 Then akuru$ = "sq": GoTo here: ctype = af                                   'seenusukun
    If lett1$ = "c" Then akuru$ = "k": ctype = ak: ' GoTo here                                                          'kaafu
    If Mid$(a$, lft, 1) = "c" And Mid$(a$, lft + 1, 1) = "h" Then akuru$ = "c": lft = lft + 1: ctype = ak: GoTo here   'chaviyani
    If lett1$ = Chr$(13) Then lines% = lines% + 1
    If Mid$(a$, lft, 1) = "o" And Mid$(a$, lft + 1, 1) = "o" Then akuru$ = "U": lft = lft + 1: ctype = f: GoTo here   'ooboofili
    If lett1$ = Mid$(a$, lft + 1, 1) Then akuru$ = "wq": ctype = af:  GoTo here                                        'alifusukun
    
    
    ' something fisshy about these 2 lines -sofwath
    
    'If InStr("aeiou. ", Mid$(a$, lft, 1)) = 0 And InStr("aeiou", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = akuru$ + "q": ctype = f:  GoTo here 'sukun
    'If InStr("aeiou.", Mid$(a$, lft, 1)) = 0 And InStr("aeiou ", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = akuru$ + "q":: GoTo here: ctype = f               'sukun
    
    
    
    'If InStr("aeiou.", Mid$(a$, lft, 1)) = 0 And InStr("aeiou", Mid$(a$, lft + 1, 1)) = 0 Then akuru$ = "q": GoTo here: 'akuru$ + "q":  ctype = f  'sukun
    
    'If Mid$(a$, lft + 1, 1) = " " And InStr("aeiou", Mid$(a$, lft - 1, 1)) <> 0 Then akuru$ = akuru$ + "q": ' GoTo here
    If Mid$(a$, lft, 1) = "." Then akuru$ = "."
    If Mid$(a$, lft, 1) = "," Then akuru$ = ","
    If Mid$(a$, lft, 1) = "'" Then akuru$ = "'"
    If Mid$(a$, lft, 1) = "!" Then akuru$ = "!" '':
    If Mid$(a$, lft, 1) = "?" Then akuru$ = "?"
    If Mid$(a$, lft, 1) = "(" Then akuru$ = ") "
    If Mid$(a$, lft, 1) = ")" Then akuru$ = "(": 'Beep
    If Mid$(a$, lft, 1) = " " Then akuru$ = " "
    If Asc(Mid$(a$, lft, 1)) = 9 Then akuru$ = " " ': ' lft = lft + 1
    
    'If Mid$(a$, lft, 1) = Chr$(13) Then akuru$ = Chr$(10) + Chr$(13)'MsgBox "ENTER"
here:
    'If lett1$ = "(" Then Beep
    If lett1$ <> " " And Mid$(a$, lft - 1, 1) = " " And Mid$(a$, lft + 1, 1) = " " Then akuru$ = Left$(akuru$, 1) ': Beep: 'GoTo here' hus akuru
        lin$ = lin$ + akuru$
    akuru$ = ""
Next
text$ = lin$
Debug.Print lin$
If ctype = f Then num_fili = num_fili + 1
End Sub
