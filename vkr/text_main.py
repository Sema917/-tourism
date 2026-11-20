# -*- coding: utf-8 -*-
"""Сборка блоков всей работы."""

from text_ch1 import chapter1
from text_ch2 import chapter2
from text_end import ending
from text_extra import missing_anchors, splice
from text_front import front_blocks
from text_intro import introduction
from text_tighten import tighten


def make_blocks(pages=None):
    blocks = []
    blocks.extend(front_blocks(pages))
    blocks.append(("break",))
    blocks.extend(introduction())
    blocks.append(("break",))
    ch1, used1 = splice(chapter1())
    blocks.extend(ch1)
    blocks.append(("break",))
    ch2, used2 = splice(chapter2())
    blocks.extend(ch2)
    missed = missing_anchors(used1 | used2)
    if missed:
        raise SystemExit("Не найдены якоря: " + " | ".join(missed))
    blocks.append(("break",))
    blocks.extend(ending())
    return tighten(blocks)
