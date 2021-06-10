import sys
import fitz
import docx2txt
from os import listdir
from os.path import isfile, join


def extract_text_from_pdf(doc_path):
    try:
        doc = fitz.open(doc_path)
        text = ""
        for page in doc:
            rect = page.rect
            words = page.get_text("words")
            mywords = [w for w in words if fitz.Rect(w[:4]).intersects(rect)]
            text = text + str(make_text(mywords))
        return text
    except KeyError():
        return ''


def extract_text_from_doc(doc_path):
    '''
    Helper function to extract plain text from .doc files

    :param doc_path: path to .doc file to be extracted
    :return: string of extracted text
    '''
    try:
        try:
            import textract
        except ImportError:
            return ' '
        text = textract.process(doc_path).decode('utf-8')
        return text
    except KeyError:
        return ' '


def extract_text_from_docx(doc_path):
    '''
    Helper function to extract plain text from .docx files

    :param doc_path: path to .docx file to be extracted
    :return: string of extracted text
    '''
    try:
        temp = docx2txt.process(doc_path)
        text = [line.replace('\t', ' ') for line in temp.split('\n') if line]
        return ' '.join(text)
    except KeyError:
        return ' '


def make_text(words):
    """Return textstring output of get_text("words").

    Word items are sorted for reading sequence left to right,
    top to bottom.
    """
    line_dict = {}
    words.sort(key=lambda w: w[0])
    for w in words:
        y1 = round(w[3], 1)
        word = w[4]
        line = line_dict.get(y1, [])
        line.append(word)
        line_dict[y1] = line
    lines = list(line_dict.items())
    lines.sort()
    return "\n".join([" ".join(line[1]) for line in lines])


allresumes = [f for f in listdir('resumes') if isfile(join('resumes', f))]
tx = None
for resume in allresumes:
    fname = 'resumes/' + resume
    extension = fname.split('.')[1]
    text = ''
    if extension == 'pdf':
        text = extract_text_from_pdf(fname)
    elif extension == 'docx':
        text = extract_text_from_docx(fname)
    elif extension == 'doc':
        text = extract_text_from_doc(fname)
    tx = " ".join(text.split('\n'))
    with open("text_resumes/" + resume.replace(".pdf", "") + ".txt", 'w', encoding='utf-8') as f:
        f.write(tx)
