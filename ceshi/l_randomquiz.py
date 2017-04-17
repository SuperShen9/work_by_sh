import random
capitals={'a':1,'b':2,'c':3,'d':4,'e':5,
          'f':6,'g':7,'h':8,'i':9,'j':10,
          'k':11,'l':12,'m':13,'n':14}

for quiznum in range(8):
    quizfile=open('capitalquiz%s.txt'%(quiznum+1),'w')
    answerkeyfile=open('capitalquiz_answer%s.txt'%(quiznum+1),'w')

    quizfile.write((' '*20)+'state capitals quiz(from %s)'%(quiznum+1))
    quizfile.write('\n\n')

    states=list(capitals.keys())
    random.shuffle(states)
    for questionnum in range(14):
        correctanswer=capitals[states[questionnum]]
        wronganswer=list(capitals.values())
        del wronganswer[wronganswer.index(correctanswer)]
        wronganswer=random.sample(wronganswer,3)
        answeroption=wronganswer+[correctanswer]
        random.shuffle(answeroption)
        quizfile.write('%s. what is the %s num?\n'%(questionnum+1,states[questionnum]))
        for i in range(4):
            quizfile.write('%s. %s\n'%('ABCD'[i],answeroption[i]))
        quizfile.write('\n')
        answerkeyfile.write(' %s. %s\n'%(questionnum+1,'ABCD'[answeroption.index(correctanswer)]))
quizfile.close()
answerkeyfile.close()

