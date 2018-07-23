import pptx

from .utils import BaseParser
from ..cors import processCors


class Parser(BaseParser):
    """Extract text from pptx file using python-pptx
    """

    def extract(self, filename, **kwargs):
        if "language" in kwargs and kwargs['language']:
            cors = processCors(kwargs["language"])
        converted_filename = filename[:-5] + '_converted.pptx'
        presentation = pptx.Presentation(filename)
        text_runs = []
        for slide in presentation.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if cors:
                            run.text = cors.apply_rules(run.text)
                            text_runs.append(run.text)
        presentation.save(converted_filename)
        return '\n\n'.join(text_runs)