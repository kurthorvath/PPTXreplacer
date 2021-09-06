from typing import List
from pptx import Presentation

import uuid
import logging
import asyncio
import threading, queue
import os
import json
import pika
import time

def search_and_replace(search_str, repl_str, input, output):
    from pptx import Presentation
    prs = Presentation(input)
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                if(shape.text.find(search_str))!=-1:
                    text_frame = shape.text_frame
                    cur_text = text_frame.paragraphs[0].runs[0].text
                    new_text = cur_text.replace(str(search_str), str(repl_str))
                    text_frame.paragraphs[0].runs[0].text = new_text
    prs.save(output)

print("start...")
search_and_replace('Franz','Herbert','IN.pptx','OUT.pptx' )
