FILE_NAME = 'Packet-Template.xlsx'

# Excel Part 
import xlrd 
import re 
wb = xlrd.open_workbook(FILE_NAME) 
sheet = wb.sheet_by_index(0) 

set_name = sheet.cell_value(0, 1) 
writers = sheet.cell_value(1, 1)

tossups = []

for i in range(21):
    answer = sheet.cell_value(3 + i, 1)
    question = sheet.cell_value(3 + i, 2)
    
    if '(*)' in question:
        question = '\power{' + question.replace('(*)", "(*)}')

    # alternating (1) quotes, (2) asterisks,
    # (3) double-underscores, (4) single-underscores
    
    question = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",question))))            
    answer = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",answer))))            

    tossups.append({
                'question': question,
                'answer': answer
            })
    
bonuses = []

for i in range(21):
    lead_in   = sheet.cell_value(25 + i, 1)
    question1 = sheet.cell_value(25 + i, 2)
    answer1   = sheet.cell_value(25 + i, 3)
    question2 = sheet.cell_value(25 + i, 4)
    answer2   = sheet.cell_value(25 + i, 5)
    question3 = sheet.cell_value(25 + i, 6)
    answer3   = sheet.cell_value(25 + i, 7)

    # alternating (1) quotes, (2) asterisks,
    # (3) double-underscores, (4) single-underscores
    
    lead_in = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",lead_in))))            
    question1 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",question1))))            
    answer1 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",answer1))))            
    question2 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",question2))))            
    answer2 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",answer2))))            
    question3 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",question3))))            
    answer3 = re.sub("_([^_]*)_","\\prompt{\\1}",re.sub("\_\_([^_]*)\_\_","\\\\answer{\\1}",re.sub("\*([^_]*)\*","\\\\textit{\\1}",re.sub("\[([^_]*)\]","\\pronuciationguide{\\1}",answer3))))            

    bonuses.append({
                'lead_in': lead_in,
                'question1': question1,
                'answer1': answer1,
                'question2': question2,
                'answer2': answer2,
                'question3': question3,
                'answer3': answer3,
            })

# LaTex Part

import os, subprocess

header = r'''\documentclass[]{article}
\pagestyle{plain}
\usepackage{multicol}
\usepackage[margin=0.25in]{geometry}

\setlength\parindent{0pt}

\newcounter{tossupnumber}
\newcounter{bonusnumber}

\newcommand{\power}[1]{\textbf{#1}}
\newcommand{\pronunciationguide}[1]{{\small \texttt{#1}}}
\newcommand{\answer}[1]{\textbf{\underline{#1}}}
\newcommand{\prompt}[1]{\underline{#1}}
\newcommand{\tossup}[2]{
	
	\par
	\refstepcounter{tossupnumber}
		
	\ifnum\thetossupnumber<21 
		{\thetossupnumber} 
	\else {TIEBREAKER}\fi . #1

	ANSWER: #2
	\newline
}
\newcommand{\bonus}[7]{
	
	\par
	\refstepcounter{bonusnumber}
	
	\ifnum\thebonusnumber<21 
		{\thebonusnumber} 
	\else {TIEBREAKER}\fi . #1
	
	[10] #2
	
	ANSWER: #3
	
	[10] #4
	
	ANSWER: #5
	
	[10] #6
	
	ANSWER: #7
	\newline
}

\title{''' + set_name + r'''}
\author{''' + writers + r'''}
\date{\vspace{-5ex}}

\begin{document}
	\maketitle
	
	\section*{Tossups}
	\begin{multicols*}{2}'''

footer = r'''	\end{multicols*}
\end{document}	
'''

main = ''

for tossup in tossups:
    if tossup['question'] != '':
        main = main + '\\tossup{'+ tossup['question'] +' }{'+ tossup['answer'] + '}'    
    
main = main + '\n \\newpage \\section*{Bonuses} \n'

for bonus in bonuses:
    if bonus['question1'] != '':
        main = main + '\\bonus{' + bonus['lead_in'] + '}{' + bonus['question1'] + '}{' + bonus['answer1'] + '}{' + bonus['question2'] + '}{' + bonus['answer2'] + '}{' + bonus['question3'] + '}{' + bonus['answer3'] + '}'

content = header + main + footer

filename = set_name.replace(' ', '_')
#os.unlink(filename + '.pdf')
with open(filename + '.tex','w') as f:
     f.write(content)

commandLine = subprocess.Popen(['pdflatex', filename + '.tex'])
commandLine.communicate()

os.unlink(filename + '.aux')
os.unlink(filename + '.log')
os.unlink(filename + '.tex')
