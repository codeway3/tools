import re


f = open("path", "r")
newf = open("newpath", "w")
for line in f:
    newline = re.sub(r' uuid="\S*"', '', line)
    newf.write(newline)
newf.close()
f.close()
