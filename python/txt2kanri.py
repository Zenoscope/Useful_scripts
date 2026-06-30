# It also needs the proper format for the comment field though.
# convert a text file to json so it can be imported into kanri

"""
Import script:
Loop through each line:

* is header, don't change until new line
- is topic, so write $header + $topic
_ are notes, so add to notes field (needs to have a topic).

__Future features:
Read input from a file
Multiple lines from notes (break on blank line or *)
Create a new file to import

"""

import random, string
def randomString():
    return ''.join(random.choice(string.ascii_lowercase + string.digits) for _ in range(25))

#
#Example
#
text="""* Recipe Vault
- change icon
_Upload to phone

*kanji study
- Compare readings and meanings
- Grid of readings/meanings, like words which have the kanji
- Mnemonic image of meaning
_ Strokes
"""


entry = "start"
header, comment, title = "","",""

counter = 0
lines = text.splitlines()

for index, line in enumerate(lines):
# for line in text.splitlines():

    if "*" in line[:1]:
        header = line
    elif "-" in line[:1]:
        title = line

        if index + 1 < len(lines):
            peek_next = lines[index + 1]
            if "_" in peek_next[:1]:
                # loop until the line is a space or *
                comment = peek_next
        entry = "end"

    if entry == "end":
        rndString = randomString()
        openBracket = "{ \n"
        idNum = "\"id\":\"" + rndString + "\",\n"
        title = "\"name\": \"" + header + " " + title + "\"\n},"

        if comment == "":
            commentLine = ""
        else:
            # what about multiple lines?
            commentLine = "\"description\": \"" + comment +  "\",\n"
            comment = ""
        # print(output3, output1, output2)
        print(openBracket,commentLine, idNum, title)
        entry = "start"

    counter = counter + 1
