import re

def extract_numbered_entries(blob):
    lines = blob.splitlines()
    numbered_list = []
    header = ''
    numbered = None
    buf = ''

    for line in lines:
        line = line.strip()
        if re.match('^\(?(\d+)\s*[.):].+',line): #beginning of a new numbered list
            if not numbered: #if its the first numbered list
                numbered=True
                header = buf
                buf = ''
            elif buf and numbered:
                numbered_list.append(buf) #push the last numbered list
                buf = '' #empty buffer
                numbered=True
        #continue buffering
        buf += (line +'\n')

    #closing block
    if buf and numbered: #if there is a buffer and its a numbered list
        numbered_list.append(buf) #push the last numbered list
        buf = ''

    return header,numbered_list
