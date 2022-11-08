from pdf2image import convert_from_path
from pptx import Presentation
from pptx.util import Pt
import os

class Slides:
    _slides = None
    _load = False
    _path = None
    
    def __init__(self, file):
        self._path, ext = os.path.splitext(file)
        if (ext == ".pdf"):
            self._slides = convert_from_path(file)
            self._load = True
        else:
            print("This file is not PDF file.")

    def toPPTX(self):
        if (self._load):
            presentation = Presentation()
            presentation.slide_height = Pt(self._slides[0].height)
            presentation.slide_width = Pt(self._slides[0].width)
            for slide in self._slides:
                pptx_slide = presentation.slides.add_slide(presentation.slide_layouts[6])
                path = self._path + "_page.png"
                slide.save(path, "png")
                pptx_slide.shapes.add_picture(path, 0, 0)
            os.remove(path)
            presentation.save(self._path+".pptx")

        else:
            print("Could not load File.")


def main():
    slides = Slides("./files/test.pdf")
    slides.toPPTX()

if __name__ == "__main__":
    main()