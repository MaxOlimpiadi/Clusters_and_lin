# -*- coding: utf-8 -*-
"""
Created on Thu Apr 24 15:02:00 2025

@author: zidan
"""
import json


#TODO: check one more time

def extract_decoded_span(file_path, span):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            raw = f.read()

        # Преобразуем escape-последовательности в нормальные символы
        decoded = codecs.decode(raw, 'unicode_escape')

        start, end = span
        if 0 <= start < end <= len(decoded):
            print("Фрагмент текста:")
            print(decoded[start:end])
        else:
            print(f"Интервал [{start}, {end}] вне диапазона (0–{len(decoded)})")
    except Exception as e:
        print(f"Ошибка: {e}")
        

data = json.load(open("Die Judenbuche.json"))
# for sent in data["sentences"]:
#     print(f'{data["text"][sent[0]:sent[1]]} \n')

start = int(input("Enter start of the span: "))
end = int(input("Enter end of the span: "))
print(data["text"][start:end])


# i = 1 
# for sentence in data["sentences"][0:100]:
#     print(f'\n Sentence #{i}: \n  {text[sentence[0]:sentence[1]]}')
#     i += 1


# for i, ch in enumerate(text_clean[:200]):
#     print(f"{i:>3}: {repr(ch)}")


