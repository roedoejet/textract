from .utils import BaseParser
from ..cors import processCors
import chardet

class Parser(BaseParser):
    """Parse ``.txt`` files"""

    def extract(self, filename, **kwargs):
        
        with open(filename, 'r') as stream:
            print(chardet.detect(stream))
            text = stream.read()
        if "language" in kwargs and kwargs['language']:
            self.cors = processCors(kwargs["language"])
            text = self.cors.apply_rules(text)
                        
        return text
