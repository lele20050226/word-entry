#!/usr/bin/env python

import sys
import urllib2
import re
print('hello')
dict_web = 'http://www.dict.cn/'

def main():
    try:
        if len(sys.argv) == 2:
            word_lookup = sys.argv[1]
        else:
            print('What do you want to look up?')
            sys.exit(1)

        f = urllib2.urlopen(dict_web+word_lookup)
        result = f.read()
        temp_str = re.search(r'<ul class="dict-basic-ul">(?P<test>.*?)</ul>', result, re.M|re.I|re.S|re.U).group('test').decode('utf8')
        print re.sub(r'<[/a-z]*>|\s\s+','',temp_str)
    except AttributeError:
        print "word can not be found"
        sys.exit(2)

    temp_str = re.search(r'<ol slider="2">(?P<test>.*?)</ol>', result, re.M|re.I|re.S|re.U).group('test').decode('utf8')
    lines = re.findall(r'<li>(?P<orig>.*?)</li>', temp_str, re.M|re.I|re.S|re.U)
    for i in range(len(lines)):
        print "%d. "%(i+1)
        print re.sub(r'<[/em].*?>|<[/li].*?>|<br/>|\s\s+','',lines[i])


if __name__ == '__main__':
	main()
