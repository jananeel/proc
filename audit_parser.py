import re

#
# This class is used to parse the numbered entries, observations etc..
# It basically extracts out the numbered entries in to an array, separating out headers and observations.
# Its still not generic enough as an interface to be standardized. But works currently

# need to move the whole parsing logic in to this class. eventually should become an interface like:
# ap = AuditParser(finding, finding_detail1,finding_detail2)
# finding_pairs = ap.getRows()
# for (finding,fdetail) in finding_pairs:
#    write_row(finding,fdetail)

class Context:
    hdr = 0,    #header - any text before a numbered list
    numbered = 1, #numbered - admiss a numbered list
    new_number = 2, #new_numbered - start of a new numbered list
    observation = 3, #observation
    end = 4 #end of text


class AuditParser:

    def __init__(self,blob):
        self.blob = blob
        self.numbered_list = []
        self.state = Context.hdr
        self.observation = ''
        self.header = ''
        self.buf = ''

    def transition(self,old_state,new_state):  #this method is called only when a transition from one context to another
        if old_state == Context.numbered:
            self.numbered_list.append(self.buf)
        elif old_state == Context.hdr:
            self.header = self.buf
        elif old_state == Context.observation:
            self.observation = self.buf
            
        self.buf = ''

        self.state = new_state
        if(new_state == Context.new_number):
            self.state = Context.numbered 

    def parse(self):
        for line in self.blob.splitlines():
            line = line.strip()
            cur_state = new_state = self.state 
            if re.match('^\(?(\d+)\s*[ .):].+',line): #beginning of a new numbered list
                new_state = Context.new_number
            elif re.match('^observation(s)?:',line,re.I):
                new_state = Context.observation
            
            if cur_state != new_state:
                self.transition(cur_state,new_state)
                
            #continue buffering
            self.buf += (line +'\n')

        self.transition(self.state,Context.end)
                    
        return self.header,self.numbered_list,self.observation

            

