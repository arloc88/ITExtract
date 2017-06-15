from docx import Document
import readDocx
import re
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

#regolar expression usata
regex = r"(?s)((if\s).*?(then).+?(?=else|if)|(else).+?(\n))"

root = Tk(className="ITExtract")
#foo = Label(root, text=" Benvenuto in ITExtract", ) # add a label to root window
root.geometry("640x640")

fname=Canvas (height=400, width=600)
fname.pack(side=TOP)
logo = PhotoImage(file="Itextractlogo.jpg")
image = fname.create_image(300,200,anchor=CENTER, image=logo)
fname.pack()

def act(): # defines an event function - for click of button
    filedocx = filedialog.askopenfilename()

    # trasformo il testo da docx a testo trattabile e applico la regex
    test_str = readDocx.getText(filedocx)
    lenght = len(test_str)
    print('Il documento analizzato contiene %d caratteri' % lenght)
    matches = re.finditer(regex, test_str, re.IGNORECASE | re.VERBOSE | re.MULTILINE)

    # costruttore docx
    doc = Document()
    doc.add_heading('Result', 0)
    size = 0

    # ciclo for che tramite regex mi individua il costrutto if then else
    for matchNum, match in enumerate(matches):
        matchNum = matchNum + 1

        m = ("\n ********************** Match {matchNum} was found at {start}-{end} **********************: "
             "\n {match}".format(matchNum=matchNum, start=match.start(), end=match.end(), match=match.group()))
        # print(m.encode("utf-8"))

        # calcolo dell'estrazione parziale e aggiornamento dell'estrazione totale
        lengroup = (match.end() - match.start())
        print("\n Match lungo %d caratteri" % lengroup)
        size = size + lengroup
        print("\n In totale sono stati estratti %d caratteri " % size)
        percent = size / lenght * 100
        print("\n Percentuale di documento analizzata %s per cento" % round(percent, 2))

        # scrittura su file docx di output
        doc.add_paragraph(m)
        # font = p.add_run(match.group()).font

    # file di output
    doc.add_heading('Stats', 0)
    doc.add_paragraph('Il documento analizzato contiene %d caratteri' % (len(test_str)))
    doc.add_paragraph('In totale sono stati estratti %d caratteri' % size)
    doc.add_paragraph('Percentuale del documento trattata: %s ' % round(percent, 2))

    messagebox.showwarning(message='Analisi completata con successo, ora scegli il file di output')
    filename = filedialog.asksaveasfilename(defaultextension=" .docx ", filetypes=[('word files', ' .docx ')])
    doc.save(filename)
    messagebox.showinfo(message='Salvataggio effettuato')

b = Button(None,text="Scegli il file di Input",command=act,) # create & configure widget 'button"
b.place(relx=0.5, rely=0.5, anchor=CENTER)
#foo.pack()
root.mainloop()