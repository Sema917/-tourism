"""Собрать сценарии, презентации и тетрадь."""

from build_handout import build as build_handout
from build_pptx import main as build_pptx
from build_scenarios import build_all
from ddo import validate


def main():
    validate()
    build_pptx()
    handout = build_handout()
    scenarios = build_all()
    print(handout)
    for path in scenarios:
        print(path)


if __name__ == "__main__":
    main()
