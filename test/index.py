import openpyxl
from openpyxl.cell.cell import VALID_TYPES
from botok.modifytokens.splitaffixed import split_affixed
from pybo.utils.regex_batch_apply import batch_apply_regex, get_regex_pairs
from botok.tokenizers.wordtokenizer import WordTokenizer
import csv
from pathlib import Path


def tokenize_line(line, wt, rules):
    """tokenize word from line

    Args:
        line (str): line from a para
        wt (obj): word tokenizer objet

    Returns:
        str: tokenized line
    """

    new_line = ''
    tokens = wt.tokenize(line, split_affixes=True)

    for token in tokens:
        new_line += f'{token.text} '
    new_line = new_line.strip()
    normalized_line = normalize_line(new_line, rules)

    return normalized_line


def normalize_line(line, rules):
    normalized_line = batch_apply_regex(line, rules)
    return normalized_line


def tokenize_text(text, wt,rules):
    new_text = ''
    lines = text.splitlines()
    for line in lines:
        new_text += tokenize_line(line, wt,rules) + '\n'
    return new_text


if __name__ == "__main__":
    wt = WordTokenizer()
    regex_file = Path('./regex.txt')
    rules = get_regex_pairs(regex_file.open(encoding="utf-8-sig").readlines())
    tokenized_text = ''
    text_path = Path('./test/bo_text.txt/')
    text = text_path.read_text(encoding='utf8')
    tokenized_text += tokenize_text(text, wt, rules)
    Path('tokenized_bo.txt').write_text(tokenized_text)
