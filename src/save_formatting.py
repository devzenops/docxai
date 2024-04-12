from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE


def copy_style(run):
    attr_dict = dict()
    for attr in dir(run.font):
        if not attr.startswith('_') and attr != 'part':
            attr_dict[attr] = getattr(run.font, attr)

    return attr_dict


def apply_style(run, style):
    for attr in style.keys():
        if attr not in ['color', 'element']:
            setattr(run.font, attr, style[attr])