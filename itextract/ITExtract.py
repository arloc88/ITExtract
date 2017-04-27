import re
import readDocx

regex = r"""
	(?s)((\sif).*?(then|\,).+?(?=else|\n)|(else).+?(\.))
	"""
test_str = readDocx.getText('prova.docx')


matches = re.finditer(regex, test_str, re.IGNORECASE | re.VERBOSE | re.MULTILINE)

for matchNum, match in enumerate(matches):
    matchNum = matchNum + 1

    print ("\n ********************** Match {matchNum} was found at {start}-{end} **********************: \n {match}".format(matchNum=matchNum, start=match.start(),
                                                                         end=match.end(), match=match.group()))

    for groupNum in range(0, len(match.groups())):
        groupNum = groupNum + 1

        print (" \n  \t ==>info: Group {groupNum} found at {start}-{end}:'{group}'".format(groupNum=groupNum, start=match.start(groupNum),
                                                                         end=match.end(groupNum),
                                                                         group=match.group(groupNum)))