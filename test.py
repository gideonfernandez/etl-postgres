import re

def findWholeWord(w):
    return re.compile(r'\b({0})\b'.format(w), flags=re.IGNORECASE).search

print(findWholeWord('test')('em TEST delete'))    # -> <match object>
x = findWholeWord('test')('AAFP/Research Hub Testimonial Creation (Op-ed)')

print('x is:', type(x))