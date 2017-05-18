import re
import readDocx

'''
regex2 = r"""
	(?s)((\sif).*?(then).+?(?=else|\n)|(else).+?(\n))
	"""
'''
regex = r"(?s)((if\s).*?(then).+?(?=else|if)|(else).+?(\n))"

test_str = readDocx.getText('prova.docx')


print('Il documento analizzato contiene %d caratteri' % (len(test_str)))


matches = re.finditer(regex, test_str, re.IGNORECASE | re.VERBOSE | re.MULTILINE)
size=0

for matchNum, match in enumerate(matches):
    matchNum = matchNum + 1

    print ("\n ********************** Match {matchNum} was found at {start}-{end} **********************: \n {match}".format(matchNum=matchNum, start=match.start(),

                                                                   end=match.end(), match=match.group()))
    lengroup = (match.end() - match.start())
    print("\n Match lungo %d caratteri" % lengroup)
    size= size+lengroup
    print("\n In totale sono stati estratti %d caratteri " %size)


    '''
    for groupNum in range(0, len(match.groups())):
        groupNum = groupNum + 1

        print (" \n  \t ==>info: Group {groupNum} found at {start}-{end}:'{group}'".format(groupNum=groupNum, start=match.start(groupNum),
                                                                         end=match.end(groupNum),
                                                                         group=match.group(groupNum)))
'''



