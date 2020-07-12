from io import StringIO
from html.parser import HTMLParser

class HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()
        self.startTags = []
    def handle_data(self, d):
        self.text.write(d)
    def get_data(self):
        return self.text.getvalue()

    def handle_starttag(self, tag, attrs):
        return super().handle_starttag(tag, attrs)

    def handle_startendtag(self, tag, attrs):
        kk =super().handle_startendtag(tag, attrs)   
        return kk

    """description of class"""


#def strip_tags(html):
#    s = HTMLStripper()
#    s.feed(html)
#    return s.get_data()


