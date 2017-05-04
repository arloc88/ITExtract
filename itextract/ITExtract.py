import re
import readDocx
'''
regex2 = r"""
	(?s)((\sif).*?(then).+?(?=else|\n)|(else).+?(\n))
	"""
'''
regex = r"(?s)((if\s).*?(then).+?(?=else|if)|(else).+?(\n))"
test_str = readDocx.getText('prova.docx')


matches = re.finditer(regex, test_str, re.IGNORECASE | re.VERBOSE | re.MULTILINE)

for matchNum, match in enumerate(matches):
    matchNum = matchNum + 1

    print ("\n ********************** Match {matchNum} was found at {start}-{end} **********************: \n {match}".format(matchNum=matchNum, start=match.start(),
                                                                       end=match.end(), match=match.group()))
    '''# Scrivi risultati in un file .
    matches = open("result.txt", "w")
    matches.write("Risultati estrazione\nLook at it and see\n \n ********************** Match {matchNum} was found at {start}-{end} **********************: \n {match}".format(matchNum=matchNum, start=match.start(),
                                                                   end=match.end(), match=match.group()))
    matches.close()
'''
    for groupNum in range(0, len(match.groups())):
        groupNum = groupNum + 1

        '''print (" \n  \t ==>info: Group {groupNum} found at {start}-{end}:'{group}'".format(groupNum=groupNum, start=match.start(groupNum),
                                                                         end=match.end(groupNum),
                                                                         group=match.group(groupNum)))
        '''
