import calendar
import json
import math
import pdb

import pandas as pd
import plotly.express as px
import plotly.graph_objs as go
import plotly.io as pio
import requests

from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR

import datetime

from pr_values import Device_to_location, compliance_data_link



# Functions

def get_compliance_handwash(x):
    # get hand wash observed which has total wash duration is more than 40 seconds.
    if x['Hand Washes Observed']:
        if x['TotalWashDuration'] >= 40:
            return 1
    return 0


def get_100_compliance_handwash(x):
    #    get hand wash observed which has total wash duration is more than 40 seconds and all core steps.
    if x['Hand Washes with Total Duration >= 40 sec']:
        if x['HandWash With All Core Steps']:
            return 1
    return 0


def text_into_presentation(pres, slide_no, shape_no, para_no, run_no, words):
    """Adds text into presentation shapes
       Input: pres - name of presentation,
              slide_no - The slide to make change in.
              shape_no - The shape which contains the text.
              para_no - The paragraphs which contains the text.
              run_no - Further specifications.
              words - The actual text we want to insert.
    """
    try:
        pres.slides[slide_no].shapes[shape_no].text_frame.paragraphs[para_no].runs[run_no].text = words
        print("pres.slides[", slide_no, "].shapes[", shape_no, "].text_frame.paragraphs[", para_no, "].runs[", run_no,
              "].text = ", words, ")", "has run correctly.")
    except:
        print("pres.slides[", slide_no, "].shapes[", shape_no, "].text_frame.paragraphs[", para_no, "].runs[", run_no,
              "].text = ", words, ")", "has not run correctly.")


def table_cell_into_presentation(pres, slide_no, shape_no, cell_x, cell_y, words):
    """Adds text into the cells of the presentation tables.
       Input: pres - name of presentation,
              slide_no - The slide to make change in.
              shape_no - The shape which contains the text.
              cell_x - The position of the cell along the x-axis.
              cell_y - The position of the cell along the y-axis.
              words - The actual text we want to insert.
    """
    try:
        pres.slides[slide_no].shapes[shape_no].table.cell(cell_x, cell_y).text = words
        print("pres.slides[", slide_no, "].shapes[", shape_no, "].table.cell(", cell_x, ", ", cell_y,
              ").text = ", words, ")", "has run correctly.")
    except:
        print("pres.slides[", slide_no, "].shapes[", shape_no, "].table.cell(", cell_x, ", ", cell_y,
              ").text = ", words, ")", "has not run correctly.")


def img_into_presentation(pres, slide_no, path, pos_x, pos_y, size_x, size_y):
    """Inserts an image into the presentation.
       Input: pres - name of presentation,
              slide_no - The slide to make change in.
              path - The path of the image.
              pos_x - The position of the image along the x-axis.
              pos_y - The position of the image along the y-axis.
              size_x - The width of the image.
              size_y - The length of the image.
    """
    try:
        pres.slides[slide_no].shapes.add_picture(path, Inches(pos_x), Inches(pos_y), width=Inches(size_x), height=Inches(size_y))
        print("pres.slides[", slide_no, "].shapes.add_picture(", path, ", Inches(", pos_x, "), Inches(", pos_y, "), "
                                                                                                                "width= Inches(",
              size_x, ")", "has run correctly.")
    except:
        print("pres.slides[", slide_no, "].shapes.add_picture(", path, ", Inches(", pos_x, "), Inches(", pos_y, "), "
                                                                                                                "width= Inches(",
              size_x, ")", "has run correctly.")


def add_text_to_shape(pr_ppt, slide_no, shape_no, text, color, most_common, left, top, width, height):
    # Get the slide with the existing shape number
    slide = pr_ppt.slides[slide_no]
    #existing_shape = slide.shapes[shape_no]

    #left, top, width, height = existing_shape.left, existing_shape.top, existing_shape.width, existing_shape.height
    print(shape_no, left, top, width, height)
    # remove the existing shape
    #sp = existing_shape._element
    #sp.getparent().remove(sp)

    # create new shape with text
    new_shape = slide.shapes.add_textbox(left, top, width, height)
    new_shape.text_frame.text = text
    new_shape.text_frame.word_wrap = True
    new_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(*color)
    new_shape.text_frame.paragraphs[1].font.color.rgb = RGBColor(*color)

    # set font properties
    if most_common:
        new_shape.text_frame.paragraphs[0].font.bold = True
        try:
            new_shape.text_frame.paragraphs[3].font.color.rgb = RGBColor(*color)
            new_shape.text_frame.paragraphs[3].font.bold = True
        except IndexError:
            pass
        
        for paragraph in new_shape.text_frame.paragraphs:
            paragraph.font.name = 'Poppins Light'
            paragraph.font.size = Pt(11)
    else: 
        for paragraph in new_shape.text_frame.paragraphs:
            paragraph.font.name = 'Poppins Light'
            paragraph.font.size = Pt(8)

        # limit the length of each line to 30 characters
        words = paragraph.text.split()
        for i in range(1, len(words)):
            if len(words[i-1]) + len(words[i]) > 30:
                words[i-1] += '\n'
        paragraph.text = ' '.join(words)



def add_text_to_shape2(pr_ppt, slide_no, text, color, left, top, width, height, rotation):
    # Get the slide with the existing shape number
    slide = pr_ppt.slides[slide_no]
    #existing_shape = slide.shapes[shape_no]

    #left, top, width, height = existing_shape.left, existing_shape.top, existing_shape.width, existing_shape.height
    print(text, left, top, width, height, rotation)
    # remove the existing shape
    #sp = existing_shape._element
    #sp.getparent().remove(sp)

    # create new shape with text
    new_shape = slide.shapes.add_textbox(left, top, width, height)
    new_shape.text_frame.text = text
    #new_shape.text_frame.word_wrap = True
    new_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(*color)
    #new_shape.text_frame.paragraphs[1].font.color.rgb = RGBColor(*color)
    new_shape.text_frame.paragraphs[0].font.name = 'Poppins Light'
    new_shape.text_frame.paragraphs[0].font.size = Pt(8)
    new_shape.rotation = rotation




def largest_remainder_rounding(numbers, target_sum):
    # Sort the numbers by their decimal parts in decreasing order
    sorted_numbers = sorted(numbers, key=lambda x: x % 1, reverse=True)
    print(sorted_numbers)
    # Round down all the numbers and compute their sum
    floor_sum = floor_sum = sum(round(n*100) for n in numbers)
    print(floor_sum)
    # Compute the difference between the target sum and the floor sum
    difference = target_sum - floor_sum
    print(difference)
    # Distribute the remaining difference by adding 1 to the largest decimal parts
    for i in range(difference):
        sorted_numbers[i] = int(sorted_numbers[i]) + 1

    return sorted_numbers

numbers = [0.13626332*100, 0.47989636*100, 0.09596008, 0.28788024]
target_sum = 100

rounded_numbers = largest_remainder_rounding(numbers, target_sum)
print(rounded_numbers)