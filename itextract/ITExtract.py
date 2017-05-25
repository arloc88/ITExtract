from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import readDocx
import re

# regolar expression usata
regex = r"(?s)((if\s).*?(then).+?(?=else|if)|(else).+?(\n))"

# nome del file su cu si vuole applicare l'estrazione
filedocx = 'prova.docx'

# trasformo il testo da docx a testo trattabile e applico la regex
test_str = readDocx.getText(filedocx)
print('Il documento analizzato contiene %d caratteri' % (len(test_str)))
matches = re.finditer(regex, test_str, re.IGNORECASE | re.VERBOSE | re.MULTILINE)

# costruttore docx
doc = Document(filedocx)

size = 0
# ciclo for che tramite regex mi individua il costrutto if then else
for matchNum, match in enumerate(matches):
    matchNum = matchNum + 1

    print("\n ********************** Match {matchNum} was found at {start}-{end} **********************: "
          "\n {match}".format(matchNum=matchNum, start=match.start(), end=match.end(), match=match.group()))

    # ogni estrazione viene evidenziata ed aggiunta al file originario
    # search_word = match.group()
    p = doc.add_paragraph(match.group())
    font = p.add_run(match.group()).font
    font.highlight_color = WD_COLOR_INDEX.YELLOW
    # print(test_str.replace(search_word, '\033[44;33m{}\033[m'.format(search_word)))

    # calcolo dell'estrazione parziale e aggiornamento dell'estrazione totale
    lengroup = (match.end() - match.start())
    print("\n Match lungo %d caratteri" % lengroup)
    size = size + lengroup
    print("\n In totale sono stati estratti %d caratteri " % size)

# file di output
doc.save('Highlight.docx')

'''
    #informazione sui match e sui gruppi
    for groupNum in range(0, len(match.groups())):
        groupNum = groupNum + 1

        print (" \n  \t ==>info: Group {groupNum} found at {start}-{end}:'{group}'".format(groupNum=groupNum, 
        start=match.start(groupNum), end=match.end(groupNum), group=match.group(groupNum)))
'''
