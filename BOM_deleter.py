import os
import codecs

BUFSIZE = 4096
BOMLEN = len(codecs.BOM_UTF8)

with open("path", "r+b") as fp:
    chunk = fp.read(BUFSIZE)
    if chunk.startswith(codecs.BOM_UTF8):
        i = 0
        chunk = chunk[BOMLEN:]
        while chunk:
            fp.seek(i)
            fp.write(chunk)
            i += len(chunk)
            fp.seek(BOMLEN, os.SEEK_CUR)
            chunk = fp.read(BUFSIZE)
        fp.seek(-BOMLEN, os.SEEK_CUR)
        fp.truncate()
