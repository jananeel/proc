from audit_parser import *

simple_numbered_string = """1. This is first numbered
2. This is second numbered
3. This is third

"""
apn = AuditParser(simple_numbered_string)
(hdr,numbered_list,obser) = apn.parse()
assert hdr == ""
assert len(numbered_list) == 3
assert len(numbered_list[2].splitlines()) == 2
assert obser == ""

header_numbered_string = """ Header is some random stuff
multilined
that is present before a numbered entry 1. It can also have numbers
1. Whatever
2. said and done
1 This is taken care of
(1) This is a numbered list too
3: So am I
observation What is this?
A multilined observation?
"""

apho = AuditParser(header_numbered_string)
(hdr,numbered_list,obser) = apho.parse()
assert hdr != '' and len(hdr.splitlines()) == 3
assert len(numbered_list) == 5
assert len(obser.splitlines()) == 2

(hdr,numbered_list,obser) = AuditParser(header_numbered_string).parse(None)
assert hdr != '' and len(hdr.splitlines()) == 3
assert len(numbered_list) == 5
assert obser == ''
 
header_observ_string = """This is just some random text with 
observations: 
Does this have a meaning to it?
"""
aph = AuditParser(header_observ_string)
(hdr,numbered_list,obser) = aph.parse()
assert hdr != '' and len(hdr.splitlines()) == 1
assert len(numbered_list) == 0
assert len(obser.splitlines()) == 2

(hdr,numbered_list,obser) = AuditParser(header_observ_string).parse(None)
assert hdr != '' and len(hdr.splitlines()) == 3
assert len(numbered_list) == 0
assert obser == ''

header_nobserv_string = """ Header is some random stuff
multilined
that is present before a numbered entry 1. It can also have numbers
1. Whatever
2. said and done
1 This is taken care of
(1) This is a numbered list too
3: So am I
4.observations: What is this?
A multilined observation and numbered?
"""
(hdr,numbered_list,obser) = AuditParser(header_nobserv_string).parse()
assert hdr != '' and len(hdr.splitlines()) == 3
assert len(numbered_list) == 5
assert len(obser.splitlines()) == 2

(hdr,numbered_list,obser) = AuditParser(header_nobserv_string).parse(None)
assert hdr != '' and len(hdr.splitlines()) == 3
assert len(numbered_list) == 6
assert obser == ''

print " All test cases passed!"
