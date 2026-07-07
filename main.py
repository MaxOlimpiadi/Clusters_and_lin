# -*- coding: utf-8 -*-
"""
Created on Wed Jun 18 11:26:45 2025

@author: zidan
"""
import codecs
import json
import math
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment

import os

import random
import shutil
from sklearn.model_selection import train_test_split
import glob
import copy

from transformers import MT5Tokenizer

#combining annotation tags by authors (converting into a more convenient format)
#--- full_annotations - the resulting data frame of updated format
#--- output.json - file, where formated data is being saved
def combining_annotations_by_authors(data):
    full_annotations = {}
    for author in data["participations"]:
        full_annotations[author] = {
            "mentions": data.get("mentions", {}).get(author, {}),
            "participations": data["participations"][author]
        }
    full_annotations["title"] = data["title"]
    full_annotations["text"] = data["text"]
    full_annotations["sentences"] = data["sentences"]
    full_annotations["tokens"] = data["tokens"]
    with open("output.json", "w", encoding="utf-8") as f:
        json.dump(full_annotations, f, ensure_ascii=False, indent=2)     
        
    return full_annotations


#OLD VERSION OF THE FUNCTION! getting mention span associated with particular participation spans and character:
#--it could be more than 1 mention span!
# def get_mention_spans(participation_spans, full_annotations, author, pers, min_overlap = 2):
    
#     result_spans = []
        
#     # Найдём минимальный старт и максимальный конец всех participation spans
#     min_part_start = min(span[0] for span in participation_spans)
#     max_part_end = max(span[1] for span in participation_spans) 

#     for mention in full_annotations[author]["mentions"].get(pers, []):
#         mention_spans = mention["spans"]

#         # Если все mention_spans начинаются после конца participation_spans — прерываем цикл
#         if all(start > max_part_end for (start, _) in mention_spans):
#             break  # упорядочены — дальше пересечений не будет

#         # Если все mention_spans заканчиваются до начала participation — пропускаем
#         if all(end < min_part_start for (_, end) in mention_spans):
#             continue

#         for a, b in mention_spans: #a, b - пара чисел (т.е. спан). И так для каждого спана из mention_spans
#             for c, d in participation_spans:
#                 if b < c:
#                     break  # так как упорядочены, дальше смысла проверять нет
#                 overlap_start = max(a, c)
#                 overlap_end = min(b, d)
#                 if overlap_end - overlap_start >= min_overlap: #значит пересеклись
#                     #result_spans.append([a, b]) - было так. Типа возвращали мэншэн целиком.
#                     result_spans.append([overlap_start, overlap_end]) # возвращаем именно само пересечение, а не целиком мэншэн
#                     break #чтобы избежать многократного добавления одного и того жe спана [a, b]
    
#     #++++++++++++++=19.11.2025 допускаем что mention участника мог быть не в самой фразе, а до неё:
#     if not result_spans:    # значит перс в партисипейшен был, но его меншен во фразе не найден. Значит, берём первый ближайший меншен слева от фразы
#         print(f'For {participation_spans} there is a mention of {pers} outside the corresponding phrase!')
#         is_mention_outside = True
        
#     else:
#         is_mention_outside = False                  
    
    
#     return result_spans, is_mention_outside #if there is no intersection - returning an empty list.




def get_mention_spans(participation_spans, full_annotations, author, pers, min_overlap=2):
    """
    participation_spans: list of [start, end] for текущего participation
    full_annotations: полный json с аннотациями
    author: 'gold' / 'author1' / ...
    pers: имя персонажа (ключ в mentions)
    min_overlap: минимальная длина пересечения, чтобы считать, что mention лежит внутри фразы
    """
    
    
    result_spans = []

    # safety
    if not participation_spans:
        return [], True

    # общий диапазон фразы (participation)
    min_part_start = min(span[0] for span in participation_spans)
    max_part_end   = max(span[1] for span in participation_spans)

    mentions = full_annotations[author]["mentions"].get(pers, [])

    # кандидаты для "ближайшего слева" mention'а
    nearest_left_mention_spans = None
    nearest_left_dist = None  # расстояние от конца mention до начала participation

    for mention in mentions:
        mention_spans = mention["spans"]

        # --- 1) кандидаты слева (для fallback) ---
        # mention целиком слева от participation (все его спаны заканчиваются раньше начала фразы)
        if all(end < min_part_start for (_, end) in mention_spans):
            # берём правый (максимальный по end) спан этого mention'а, чтобы измерить близость
            rightmost_span = max(mention_spans, key=lambda s: s[1])
            dist = min_part_start - rightmost_span[1]
            if nearest_left_dist is None or dist < nearest_left_dist:
                nearest_left_dist = dist
                nearest_left_mention_spans = mention_spans  # сохраним все спаны этого mention'а
            # этот mention гарантированно не пересекается, переходим к следующему
            continue

        # --- 2) если mention целиком справа от participation, дальше можно прервать цикл ---
        # (предполагаем, что mentions отсортированы по тексту)
        if all(start > max_part_end for (start, _) in mention_spans):
            break

        # --- 3) поиск реального пересечения mention'а с participation ---
        for a, b in mention_spans:         # спаны mention'а
            for c, d in participation_spans:  # спаны participation
                overlap_start = max(a, c)
                overlap_end   = min(b, d)
                if overlap_end - overlap_start >= min_overlap:
                    # возвращаем именно пересечение, а не весь mention-спан
                    result_spans.append([overlap_start, overlap_end])
                    # break только из внутреннего цикла по participation_spans,
                    # чтобы не добавлять одинаковый [a, b] многократно
                    break

    # --- постобработка / fallback ---
    # если пересечений не нашли, допускаем, что mention мог быть вне спанов фразы
    is_mention_outside = False

    if not result_spans:
        if nearest_left_mention_spans is not None:
            # не нашли внутри фразы → используем ближайший слева mention
            # и возвращаем ВСЕ его спаны
            result_spans = nearest_left_mention_spans[:]  # shallow copy на всякий случай
            is_mention_outside = True
            # print(
            #     f'For {participation_spans} mention of {pers} is outside phrase; '
            #     f'using nearest left mention spans {nearest_left_mention_spans}'
            # )
        else:
            # персонаж есть в participation, но ни одного mention'а его имени/местоимения
            # в аннотациях (для данного author) не найдено
            is_mention_outside = True
            # print(
            #     f'For {participation_spans} there is NO mention of {pers} '
            #     f'in annotations for author "{author}"'
            # )

    return result_spans, is_mention_outside


# def get_outside_mention(participation_spans, full_annotations, author, pers, min_overlap = 2):
    
    
#     return result_spans



def get_spans_with_tags(full_annotations, author='katharina_g'):
    raw_phrase_spans = []
    raw_pers_spans = []
    phrase_group_id = 1  # Один ID на одну активность
    
    for instance in full_annotations[author]["participations"]:
        tmp_spans = instance["spans"] #all the spans of current participations
        # Добавим все спаны активности, пометив одинаковым group id
        for span_start, span_end in tmp_spans:
            raw_phrase_spans.append((span_start, span_end, "phrase", phrase_group_id))
        phrase_group_id += 1
        
        
        
        #Теперь персы:
        tmp_chars = []            #all chars of current participation (пары "перс-тип")
        #getting chars of current participation object:
        for category in ["agentive", "low_agentive", "passive"]:
            if instance[category]:
                for char in instance[category]:
                    tmp_chars.append((char, category))
  
        #обработка персов в активности:
        for char, category in tmp_chars:              
            #getting mention spans of current char in the current participation spans
            char_mention_spans = get_mention_spans(tmp_spans, full_annotations, author, char) 
            for m_span_start, m_span_end in char_mention_spans:
                raw_pers_spans.append((m_span_start, m_span_end, category, None))  # пока без id
    
    # Объединяем всё и сортируем по старту
    all_spans = raw_phrase_spans + raw_pers_spans
    all_spans.sort(key=lambda x: x[0])
    # Теперь выдаём id:
    new_spans = []
    pers_id = 1
    for start, end, tag, gid in all_spans:
        if tag == "phrase":
            new_spans.append((start, end, tag, gid))  # используем сохранённый group_id
        else:
            new_spans.append((start, end, tag, pers_id))  # нумеруем персонажей
            pers_id += 1    
    
    
    # вывод получившейся структуры:
    # for elem in new_spans:
    #     print(elem)
            
    return new_spans
                    





def get_insertions(spans_with_tags, full_annotations, author="katharina_g"):
    insertions = []
    for start, end, tag, id_num in spans_with_tags:
        insertions.append((start, f"<{tag} id=\"{id_num}\">"))
        #insertions.append((end, f"</{tag} id=\"{id_num}\">"))
        insertions.append((end, f"</{tag}>"))

    insertions.sort(key=lambda x: (x[0], 0 if '/' in x[1] else 1))
    #printing insertions:
    # print("_________________________________________________")
    # for elem in insertions:
    #     print(elem)
        
    return insertions


# def get_lin_format(full_annotations):
#     #result_text = add_phrase_tags(full_annotations)
#     spans_with_tags = get_spans_with_tags(full_annotations)




# def insert_tags_with_shift_and_save(text, insertions, output_path = 'lin_text.txt'):
#     # Сортировка вставок: сначала по позиции, затем закрывающие перед открывающими
#     insertions.sort(key=lambda x: (x[0], 0 if '/' in x[1] else 1))

#     result = []
#     shift = 0
#     last_pos = 0

#     for pos, tag in insertions:
#         adj_pos = pos + shift
#         result.append(text[last_pos:adj_pos])
#         result.append(tag)
#         last_pos = adj_pos
#         shift += len(tag)

#     result.append(text[last_pos:])
#     final_text = ''.join(result)

#     # Сохраняем в файл
#     with open(output_path, 'w', encoding='utf-8') as f:
#         f.write(final_text)

#     print(f"Размеченный текст сохранён в: {output_path}")
#     return final_text




def insert_tags_with_shift_and_save(text, insertions,  num_of_text, output_dir='lin_texts'):
    # Сортировка вставок: сначала по позиции, затем закрывающие перед открывающими
    insertions.sort(key=lambda x: (x[0], 0 if '/' in x[1] else 1))

    result = []
    shift = 0
    last_pos = 0

    for pos, tag in insertions:
        # Вставляем текст между последней позицией и текущей
        result.append(text[last_pos:pos])
        result.append(tag)
        last_pos = pos  # не pos + shift, а именно оригинальная позиция
        shift += len(tag)  # если нужно — можно использовать shift позже (например, для логов)

    result.append(text[last_pos:])
    final_text = ''.join(result)


    # Генерируем имя файла, например: lin_text_001.txt
    output_filename = f"lin_text_{num_of_text:03d}.txt"
    output_path = os.path.join(output_dir, output_filename)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_text)

    print(f"Размеченный текст сохранён в: {output_path}")
    return final_text




#определяем с какого токена по какой токен занимает данное предложение, а также все айди токенов и тексты токенов в предложении
def get_tokens_span_for_sentence(sentence_start, sentence_end, formated_full_annotations):
    #token_num = 0
    tokenIds = []
    tokenTexts = []
    for token in formated_full_annotations["tokens"]:
        if token["charBegin"] >= sentence_start and token["charEnd"] <= sentence_end:
            tokenIds.append(token["id"])
            tokenTexts.append(token["text"])
        if token["charBegin"] > sentence_end:
            break
        
    # # Begin, End (должен быть на 1 больше чем номер последнего токена), tokensIds, TokenTexts
    # return min(tokenIds), max(tokenIds) + 1, tokenIds, tokenTexts
    
    # Begin, End (должен быть на 1 больше чем номер последнего токена), tokensIds, TokenTexts
    if not tokenIds:
        #print(f"ALARM! Пустое предложение: {sentence_start}, {sentence_end}")
        return None, None, [], []
    else:
        return min(tokenIds), max(tokenIds) + 1, tokenIds, tokenTexts


# по заданному спану получаем все токены этого спана (в виде айдишников и в виде текстов)
def get_tokens_for_span(span_start, span_end, formated_full_annotations):
    tokenIds = []
    tokenTexts = []
    for token in formated_full_annotations["tokens"]:
        if token["charBegin"] >= span_start and token["charEnd"] <= span_end:
            tokenIds.append(token["id"])
            tokenTexts.append(token["text"])
        if token["charBegin"] > span_end:
            break  
    return min(tokenIds), max(tokenIds), tokenIds, tokenTexts


#готовим поле phrase под наш новый формат
# 12.08 добавил +1 к TokenEnd под Финновский формат
def get_phrase_field(participation, formated_full_annotations):
    phrase_spans = participation["spans"]
    formated_spans = []
    tokenIds = []
    tokenTexts = []
    for span_start, span_end in phrase_spans:
        tokenBegin, tokenEnd, tmpTokenIds, tmpTokenTexts = get_tokens_for_span(span_start, span_end, formated_full_annotations)
        formated_spans.append({"begin": tokenBegin, "charBegin": span_start, "charEnd": span_end, "end": tokenEnd + 1})
        tokenIds += tmpTokenIds
        tokenTexts += tmpTokenTexts
    tokenTexts = ' '.join(tokenTexts)
    return {"spans": formated_spans, "text": tokenTexts, "tokenIds": tokenIds}


#получаем поле mention (конкретного перса в конкретной активности) под наш новый формат
def get_mention_field(mention_spans, formated_full_annotations, category, char):
    formated_spans = []
    tokenIds = []
    tokenTexts = []
    for span_start, span_end in mention_spans:
        tokenBegin, tokenEnd, tmpTokenIds, tmpTokenTexts = get_tokens_for_span(span_start, span_end, formated_full_annotations)
        formated_spans.append({"begin": tokenBegin, "charBegin": span_start, "charEnd": span_end, "end": tokenEnd})
        tokenIds += tmpTokenIds
        tokenTexts += tmpTokenTexts
    tokenTexts = ' '.join(tokenTexts)
    return {"spans": formated_spans, "type": category, "character": char, "text": tokenTexts, "tokenIds": tokenIds}



#по заданному в виде [a, b] спану получаем этот спан в новом формате и дополняем итоговый текст и список адйшников для категории
#в cur_span_text - накопленный объединённый текст мэншенов,
#в cur_span_token_ids - накопленные адишник тэгов меншенов
# 12.08 добавил +1 к tokenEnd, т.к. это поле должно быть на 1 больще реального последнего айдишника, судя по json файлам Финна
def get_formated_span(span, formated_full_annotations):
    span_start, span_end = span   
    tokenBegin, tokenEnd, tmpTokenIds, tmpTokenTexts = get_tokens_for_span(span_start, span_end, formated_full_annotations)
    joined_span_text = ' '.join(tmpTokenTexts)
    
    span_obj = {
        "begin": tokenBegin,
        "charBegin": span_start,
        "charEnd": span_end,
        "end": tokenEnd + 1        
    }
     
    return span_obj, joined_span_text, tmpTokenIds




def get_category_feild_part(category, elem, full_annotations, formated_full_annotations, author = "katharina_g", special_outside_m = None, seen_spans = None):
    formated_spans = []
    cur_span_token_ids = set()
    cur_span_text = []
    elem_outside_mentions = 0
    
    if seen_spans is None:
        seen_spans = set() # если нифига не передали, типа страховка.
    
    
    for char in elem[category]:
        mention_spans, is_outside_mention = get_mention_spans(elem["spans"], full_annotations, author, char)
        if is_outside_mention:
            elem_outside_mentions += 1
            # считаем спец-кейсы "crowd" и "background_character"
            if special_outside_m is not None and char in special_outside_m:
                special_outside_m[char] += 1
                
                
        for span in mention_spans:
            span_obj, span_text, span_token_ids = get_formated_span(span, formated_full_annotations)
            
            # ключ: полностью определяет "одинаковость" спана
            span_key = (
                span_obj["charBegin"],
                span_obj["charEnd"],
                span_text,
                tuple(span_token_ids)
            )
            
            if span_key in seen_spans:
                print(f'Duplicated span: {span_obj["charBegin"]}, {span_obj["charEnd"]}, participation: {elem["spans"]}, type: {category}')
                continue
            
            seen_spans.add(span_key)
            formated_spans.append(span_obj)
            cur_span_text.append(span_text)
            cur_span_token_ids.update(span_token_ids)  # добавляем без дублей
            
            
    #сортировка нужна или не? Пока решил что нужна
    token_ids_sorted = sorted(cur_span_token_ids)
    
    return {"spans": formated_spans, "text": cur_span_text, "tokenIds": token_ids_sorted}, elem_outside_mentions
            




def format_text(full_annotations, num_of_file, author = "katharina_g"):
    #formated_full_annotations = full_annotations
    tokens_spans = full_annotations["tokens"]   #заранее сохраним токены
    sentences_spans = full_annotations["sentences"]
    
    
    
    formated_full_annotations = {"annotations": []}
    
    formated_full_annotations["documentName"] = full_annotations["title"]
    formated_full_annotations["originalText"] = full_annotations["text"]     
    formated_full_annotations["tokens"] = []     #типа будем же в новом формате токены добавлять сюда..
    formated_full_annotations["sentences"] = []
   
    #+++++++++++++++++++++++ обрабатываем ТОКЕНЫ: ++++++++++++++++++++++++++++++++++++++++++
    id_sentence = 1
    id_token = 0
    true_token_structure = []       # получаем фулл инфу о каждом из токенов
    for sentence_token_group in tokens_spans:
        if not sentence_token_group:
            continue  # пропускаем пустые предложения
        word = 1
        for token_start, token_end in sentence_token_group:
            true_token_structure.append({"charBegin": token_start, "charEnd": token_end, 
                                "id": id_token, "sentence": id_sentence,
                                "text": formated_full_annotations["originalText"][token_start:token_end],
                                "word": word
                                })
            word += 1
            id_token += 1
        id_sentence += 1
    formated_full_annotations["tokens"] = true_token_structure
    #++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    
    #+++++++++++++++++++++++ обрабатываем ПРЕДЛОЖЕНИЯ: +++++++++++++++++++++++++++++++++++
    
    id_sentence = 0
    true_sentence_structure = []
    for sentence_start, sentence_end in sentences_spans:
        tokenBegin, tokenEnd, tokenIds, tokenTexts = get_tokens_span_for_sentence(sentence_start, sentence_end, formated_full_annotations)
        
        if not tokenIds: continue
        
        true_sentence_structure.append({"begin": tokenBegin , "charBegin": sentence_start, 
                                    "charEnd": sentence_end, "end": tokenEnd,
                                    "id": id_sentence, "text": formated_full_annotations["originalText"][sentence_start:sentence_end],
                                    "tokenIds":  tokenIds, "tokens": tokenTexts})
        id_sentence += 1
    formated_full_annotations["sentences"] = true_sentence_structure
    #++++++++++++++++++++++++++++++==+++++++++++++++++++++++++++++++++++++++++++++++=+++++
    
    
    
    
    #+++++++++++++++++++++++++++++Теперь аннотации сами+++++++++++++++++++++++++++++++++++++++++++
    mentions = full_annotations["katharina_g"]["mentions"]
    participations = full_annotations["katharina_g"]["participations"]    
    pid = 0 #здесь это НЕ НОМЕР PARTICIPATION OBJECT-A, а именно номер отдельного объекта в нашем формате!
    outside_metions = 0
    special_outside_m = {"crowd": 0, "background_character": 0} # для статистики outside m по этим двум персонажам
    
    #заряжаем словарик (по категориям) уже виденных ранее спанов, чтобы дублей в рамках одного спана и одной категории избегать, как в рамках однолй фразы, так и вообще 
    global_seen_span_category = {
        "agentive": set(),
        "low_agentive": set(),
        "passive": set()
    }
    
    for elem in participations:
        #phrase_spans = elem["spans"]
        formated_mention = []
        phrase_field = get_phrase_field(elem, formated_full_annotations)
        
        agentive_field, tmp_outside_mentions = get_category_feild_part("agentive", elem, full_annotations, formated_full_annotations, special_outside_m = special_outside_m, seen_spans = global_seen_span_category["agentive"]) 
        outside_metions += tmp_outside_mentions
        low_agentive_field, tmp_outside_mentions = get_category_feild_part("low_agentive", elem, full_annotations, formated_full_annotations, special_outside_m = special_outside_m, seen_spans = global_seen_span_category["low_agentive"]) 
        outside_metions += tmp_outside_mentions
        passive_field, tmp_outside_mentions = get_category_feild_part("passive", elem, full_annotations, formated_full_annotations, special_outside_m = special_outside_m, seen_spans = global_seen_span_category["passive"]) 
        outside_metions += tmp_outside_mentions
        
        formated_full_annotations["annotations"].append({"phrase": phrase_field, 
                                                        "agentive": agentive_field,
                                                        "low_agentive": low_agentive_field,
                                                        "passive": passive_field,
                                                        "hasNested": False, 
                                                        "id": pid,
                                                        "isNested": False})

        pid += 1
    
    
    os.makedirs("formated_full_texts", exist_ok= True)
    filename = os.path.join("formated_full_texts", f'formated_output_{num_of_file}.json')
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(formated_full_annotations, f, ensure_ascii=False, indent=2)
            
    #считаем фразы с пустыми участниками:
    empty_phrases = get_phrases_with_no_participants(formated_full_annotations)
    print(f'{formated_full_annotations["documentName"]}: {empty_phrases} empty phrases')
    
    #выводим статистику меншенов за границей партисипейшенов:
    print(f'Cases with the mention outside of the participation spans: {outside_metions}')
    

    # наша новая статистика по crowd/bg
    print(
        "Special outside mentions:",
        f'crowd={special_outside_m["crowd"]}, '
        f'background_character={special_outside_m["background_character"]}'
    )   
    
    return formated_full_annotations
            
            
def delete_old_chunks():
    base_dir = os.path.abspath(".")  # Корень проекта
    chunks_dir = os.path.join(base_dir, "chunks")
    subfolders = ["dev", "train", "test"]
    
    # Удаляем файлы из папок dev, train, test
    for folder in subfolders:
        path = os.path.join(chunks_dir, folder)
        if os.path.exists(path):
            for file_path in glob.glob(os.path.join(path, "*")):
                if os.path.isfile(file_path):
                    os.remove(file_path)
                    #print(f"Удалён файл: {file_path}")
        else:
            print(f"Папка не найдена: {path}")
    
    # удаляем файлы, начинающиеся с formated_output в корне проекта
    pattern = os.path.join(base_dir, "formated_output*")
    for file_path in glob.glob(pattern):
        if os.path.isfile(file_path):
            os.remove(file_path)
            #print(f"Удалён файл: {file_path}")    

# # Валидация чанка
# def validate_chunk(cd):
#     n = len(cd["tokens"])
#     for ann in cd.get("annotations", []):
#         for key in ("phrase","agentive","low_agentive","passive"):
#             part = ann.get(key)
#             if not part: continue
#             for sp in part.get("spans", []):
#                 assert 0 <= sp["begin"] < sp["end"] <= n
#             for tid in part.get("tokenIds", []):
#                 assert 0 <= tid < n


def validate_chunk(cd, raise_on_error=False, print_ok=False):
    """
    Проверяет корректность индексов в чанке:
      - спаны: 0 <= begin < end <= len(tokens) + 1
      - tokenIds: 0 <= tid < len(tokens)

    Возвращает: (ok: bool, errors: list[str])
    Если raise_on_error=True — бросает ValueError с подробным отчётом.
    """
    name = cd.get("documentName", "<unknown>")
    tokens = cd.get("tokens", [])
    n = len(tokens)
    errors = []

    if n == 0:
        errors.append(f"[{name}] нет токенов (len(tokens)=0) — остальные проверки бессмысленны.")

    roles = ("phrase", "agentive", "low_agentive", "passive")
    #roles = ("phrase", "agentive")
    
    for a_idx, ann in enumerate(cd.get("annotations", [])):
        for role in roles:
            part = ann.get(role)
            if not part:
                continue

            # --- спаны ---
            spans = part.get("spans", [])
            for s_idx, sp in enumerate(spans):
                b = sp.get("begin")
                e = sp.get("end")
                cb = sp.get("charBegin")
                ce = sp.get("charEnd")

                if not isinstance(b, int) or not isinstance(e, int):
                    errors.append(
                        f"[{name}] ann#{a_idx} role={role} span#{s_idx}: begin/end не int "
                        f"(begin={b}, end={e})"
                    )
                    continue

                if not (0 <= b < e <= n + 1):
                    errors.append(
                        f"[{name}] ann#{a_idx} role={role} span#{s_idx}: неверный диапазон "
                        f"[begin={b}, end={e}) при n={n}; char=[{cb},{ce})"
                    )

            # --- tokenIds ---
            tids = part.get("tokenIds", [])
            for t_idx, tid in enumerate(tids):
                if not isinstance(tid, int):
                    errors.append(
                        f"[{name}] ann#{a_idx} role={role} tokenIds[{t_idx}] не int: {tid}"
                    )
                    continue
                if not (0 <= tid < n):
                    errors.append(
                        f"[{name}] ann#{a_idx} role={role} tokenIds[{t_idx}] вне диапазона: "
                        f"{tid} при n={n}"
                    )

    ok = len(errors) == 0
    if ok and print_ok:
        print(f"[{name}] OK: tokens={n}, annotations={len(cd.get('annotations', []))}")

    if (not ok) and raise_on_error:
        raise ValueError("Ошибки в чанке:\n" + "\n".join(errors))

    return ok, errors



#получаем границы текущего чанка:
def get_chunk_boundaries(chunk_data):
    chunk_sentences = chunk_data["sentences"]
    first_sentence = chunk_sentences[0]
    last_sentence = chunk_sentences[-1]
    chunk_start_token = first_sentence["begin"]
    chunk_end_token = last_sentence["end"]
    chunk_start_char = first_sentence["charBegin"]
    chunk_end_char = last_sentence["charEnd"]
    return chunk_start_token, chunk_end_token, chunk_start_char, chunk_end_char


def normalize_chunk(chunk_data, cur_comp_outside_m, cur_partly_outside_m):
    # Определяем границы чанка в токенах и в символах:
    chunk_start_token, chunk_end_token, chunk_start_char, chunk_end_char = get_chunk_boundaries(chunk_data)
    
    first_chunk_sentence_id = chunk_data["sentences"][0]["id"] # для нормализации номера предложения ("sentence") в токенах
    
    
    # if first_chunk_sentence_id == 100:
    #     print("Babka")
    #     #print(chunk_data)
    #     print(f' First sentid: {first_chunk_sentence_id}')

       
    
    # Разбираемся с предложениями:
    id_sent = 0
    for sent in chunk_data["sentences"]:
        sent["begin"] -= chunk_start_token
        sent["end"] -= chunk_start_token
        sent["charBegin"] -= chunk_start_char
        sent["charEnd"] -= chunk_start_char
        sent["id"] = id_sent
        sent["tokenIds"] = [tid - chunk_start_token for tid in sent.get("tokenIds", [])]
        id_sent += 1
    
    # Теперь с токенами:
    id_token = 0
    for token in chunk_data["tokens"]:
        token["charBegin"] -= chunk_start_char 
        token["charEnd"] -= chunk_start_char
        token["id"] = id_token
        token["sentence"] = token["sentence"] - first_chunk_sentence_id
        id_token += 1
        

        
    # И, наконец, с аннотациями:
    for ann in chunk_data["annotations"]:
        if ann["phrase"]:
            for span in ann["phrase"]["spans"]:
                span["begin"] -= chunk_start_token
                span["end"] -= chunk_start_token
                span["charBegin"] -= chunk_start_char
                span["charEnd"] -= chunk_start_char
            ann["phrase"]["tokenIds"] = [tid - chunk_start_token for tid in ann["phrase"]["tokenIds"]]
        if ann["agentive"]:
            for span in ann["agentive"]["spans"]:
                span["begin"] -= chunk_start_token
                span["end"] -= chunk_start_token
                span["charBegin"] -= chunk_start_char
                span["charEnd"] -= chunk_start_char
            
            #заново пересобираем спаны и тексты, избавляясь от невалидных спанов:    
            normalized_spans = ann["agentive"]["spans"]
            old_texts = ann["agentive"]["text"]
            true_spans = []
            true_texts = []
            for norm_span, text in zip(normalized_spans, old_texts):
                if norm_span["charBegin"] < 0:
                    if norm_span["charEnd"] < 0: # тогда спан полностью слева от чанка
                        continue
                    else:   #то есть если конец спана внутри чанка. Тогда пересечение делаем:
                        norm_span["begin"] = 0 # обрезаем слева
                        norm_span["charBegin"] = 0
                        text = chunk_data["originalText"][norm_span["charBegin"]:norm_span["charEnd"]]
                true_spans.append(norm_span)
                true_texts.append(text)
                
            ann["agentive"]["spans"] = true_spans
            ann["agentive"]["text"] = true_texts
            ann["agentive"]["tokenIds"] = [tid - chunk_start_token for tid in ann["agentive"]["tokenIds"]]
            ann["agentive"]["tokenIds"] = [tid for tid in ann["agentive"]["tokenIds"] if tid >= 0]
                         
                        
                
                
            
        if ann["low_agentive"]:
            for span in ann["low_agentive"]["spans"]:
                span["begin"] -= chunk_start_token
                span["end"] -= chunk_start_token
                span["charBegin"] -= chunk_start_char
                span["charEnd"] -= chunk_start_char
    
            #заново пересобираем спаны и тексты, избавляясь от невалидных спанов:    
            normalized_spans = ann["low_agentive"]["spans"]
            old_texts = ann["low_agentive"]["text"]
            true_spans = []
            true_texts = []
            for norm_span, text in zip(normalized_spans, old_texts):
                if norm_span["charBegin"] < 0:
                    if norm_span["charEnd"] < 0: # тогда спан полностью слева от чанка
                        continue
                    else:   #то есть если конец спана внутри чанка. Тогда пересечение делаем:
                        norm_span["begin"] = 0 # обрезаем слева
                        norm_span["charBegin"] = 0
                        text = chunk_data["originalText"][norm_span["charBegin"]:norm_span["charEnd"]]
                true_spans.append(norm_span)
                true_texts.append(text)
                
            ann["low_agentive"]["spans"] = true_spans
            ann["low_agentive"]["text"] = true_texts
            ann["low_agentive"]["tokenIds"] = [tid - chunk_start_token for tid in ann["low_agentive"]["tokenIds"]]
            ann["low_agentive"]["tokenIds"] = [tid for tid in ann["low_agentive"]["tokenIds"] if tid >= 0]
        
        
        
        
        if ann["passive"]:
            for span in ann["passive"]["spans"]:
                span["begin"] -= chunk_start_token
                span["end"] -= chunk_start_token
                span["charBegin"] -= chunk_start_char
                span["charEnd"] -= chunk_start_char
            
            #заново пересобираем спаны и тексты, избавляясь от невалидных спанов:    
            normalized_spans = ann["passive"]["spans"]
            old_texts = ann["passive"]["text"]
            true_spans = []
            true_texts = []
            for norm_span, text in zip(normalized_spans, old_texts):
                if norm_span["charBegin"] < 0:
                    if norm_span["charEnd"] < 0: # тогда спан полностью слева от чанка
                        #print(f'The outside mention was completely outside the chunk!')
                        cur_comp_outside_m += 1
                        continue
                    else:   #то есть если конец спана внутри чанка. Тогда пересечение делаем:
                        #print(f'The outside mention was partly outside the chunk!')
                        cur_partly_outside_m += 1
                        norm_span["begin"] = 0 # обрезаем слева
                        norm_span["charBegin"] = 0
                        text = chunk_data["originalText"][norm_span["charBegin"]:norm_span["charEnd"]]
                true_spans.append(norm_span)
                true_texts.append(text)
                
            ann["passive"]["spans"] = true_spans
            ann["passive"]["text"] = true_texts
            ann["passive"]["tokenIds"] = [tid - chunk_start_token for tid in ann["passive"]["tokenIds"]]
            ann["passive"]["tokenIds"] = [tid for tid in ann["passive"]["tokenIds"] if tid >= 0]
    
    return chunk_data, cur_comp_outside_m, cur_partly_outside_m            
        


def split_the_file(num_of_file):
    # === Настройки ===
    filename = f"formated_output_{num_of_file}.json"
    INPUT_FILE = os.path.join("formated_full_texts", filename)  # путь к исходному файлу JSON
    OUTPUT_DIR = "chunks"                  # папка для сохранённых блоков
    # WINDOW_SIZE = 10                     # количество предложений в одном блоке
    # STRIDE = 5 
    WINDOW_SIZE = 6                     # количество предложений в одном блоке
    STRIDE = 3                          # шаг между блоками
    DOC_COMP_OUTSIDE_M = 0 # total of completely outside mentions
    DOC_PARTLY_OUTSIDE_M = 0 # total of partly outside mentions 

    
    # === Загрузка данных ===
    with open(INPUT_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
    
    all_sentences = data.get("sentences", [])
    all_tokens = data.get("tokens", [])
    all_annotations = data.get("annotations", [])
    documentName = data.get("documentName", [])
    originalText = data.get("originalText", [])
    
    num_sentences = len(all_sentences)
    
    # Индекс для быстрого доступа по ID
    token_by_id = {token["id"]: token for token in all_tokens}
    
    # === Основной цикл с overlap ===
    for i in range(0, num_sentences, STRIDE):
        chunk_sentences = all_sentences[i:i + WINDOW_SIZE]
        if not chunk_sentences:
            break
    
        
        # Собираем ID токенов
        token_ids = set()
        for sent in chunk_sentences:
            token_ids.update(sent.get("tokenIds", []))
    
        # Извлекаем нужные токены
       #chunk_tokens = [token_by_id[tid] for tid in token_ids if tid in token_by_id]
        chunk_tokens = [token_by_id[tid] for tid in sorted(token_ids) if tid in token_by_id]
        # Отбираем аннотации, у которых есть пересечение с токенами
        # chunk_annotations = [
        #     ann for ann in all_annotations
        #     if any(tid in token_ids for tid in ann.get("phrase", {}).get("tokenIds", []))
        # ]
        
        chunk_annotations = [
            ann for ann in all_annotations
            if ann.get("phrase", {}).get("tokenIds")  # непустой список 
            and all(tid in token_ids for tid in ann.get("phrase", {}).get("tokenIds", [])) # чтобы ВСЕ токены фразы попадали в чанк. Иначе многоспановая фраза может цеплять сразу 2 чанка и вызвать ошибку
        ]
    
        # Извлекаем часть текста для текущего чанка
        char_begin = chunk_sentences[0]["charBegin"]
        char_end = chunk_sentences[-1]["charEnd"]
        chunk_text = originalText[char_begin:char_end] if originalText else ""
        
        
        # Сохраняем результат
        chunk_data = {
            "annotations": chunk_annotations,
            "documentName": f"{documentName}__{i}_{min(i+WINDOW_SIZE, num_sentences)}",
            "originalText": chunk_text,
            "sentences": chunk_sentences,
            "tokens": chunk_tokens
        }
        
        if chunk_data["annotations"] == []: 
            print(f"Empty chunk annotations! {chunk_data['documentName']}")
            continue
    
        # Нормализуем наш чанк (чтобы все токены и предложения были с нуля для чанка):
        normalized_chunk_data, DOC_COMP_OUTSIDE_M, DOC_PARTLY_OUTSIDE_M = normalize_chunk(copy.deepcopy(chunk_data), DOC_COMP_OUTSIDE_M, DOC_PARTLY_OUTSIDE_M)
        
        # Сразу проверяем на битые спаны/токены:
        ok, errs = validate_chunk(normalized_chunk_data)
        if not ok:
            # можно вывести, залогировать или сохранить в файл
            for e in errs:
                print(e)        
    
    
        filename = f"chunk_{num_of_file + 7}_OLD_KATHARINA_{i}_{min(i + WINDOW_SIZE, num_sentences)}.json"
        with open(os.path.join(OUTPUT_DIR, filename), "w", encoding="utf-8") as out_f:
            json.dump(normalized_chunk_data, out_f, ensure_ascii=False, indent=2)
    
    # === Обработка "хвоста" (последних предложений) ===
    last_start = num_sentences - (num_sentences % STRIDE)
    if last_start < num_sentences:
        chunk_sentences = all_sentences[last_start:]
        token_ids = set()
        for sent in chunk_sentences:
            token_ids.update(sent.get("tokenIds", []))
        #chunk_tokens = [token_by_id[tid] for tid in token_ids if tid in token_by_id]
        chunk_tokens = [token_by_id[tid] for tid in sorted(token_ids) if tid in token_by_id]
        # chunk_annotations = [
        #     ann for ann in all_annotations
        #     if any(tid in token_ids for tid in ann.get("phrase", {}).get("tokenIds", []))
        # ]
        chunk_annotations = [
            ann for ann in all_annotations
            if ann.get("phrase", {}).get("tokenIds")  # непустой список 
            and all(tid in token_ids for tid in ann.get("phrase", {}).get("tokenIds", [])) # чтобы ВСЕ токены фразы попадали в чанк. Иначе многоспановая фраза может цеплять сразу 2 чанка и вызвать ошибку
        ]
        
        # Извлекаем часть текста для текущего чанка
        char_begin = chunk_sentences[0]["charBegin"]
        char_end = chunk_sentences[-1]["charEnd"]
        chunk_text = originalText[char_begin:char_end] if originalText else ""
        
        chunk_data = {
            "annotations": chunk_annotations,
            "documentName": f"{documentName}__{last_start}_end",
            "originalText": chunk_text,
            "tokens": chunk_tokens,
            "sentences": chunk_sentences
        }
        
        if chunk_data["annotations"] == []: 
            print(f"Empty chunk annotations! {chunk_data['documentName']}")
        
        # Нормализуем наш чанк (чтобы все токены и предложения были с нуля для чанка):
        normalized_chunk_data, DOC_COMP_OUTSIDE_M, DOC_PARTLY_OUTSIDE_M = normalize_chunk(copy.deepcopy(chunk_data), DOC_COMP_OUTSIDE_M, DOC_PARTLY_OUTSIDE_M)
        
        # Сразу проверяем на битые спаны/токены:
        ok, errs = validate_chunk(normalized_chunk_data)
        if not ok:
            # можно вывести, залогировать или сохранить в файл
            for e in errs:
                print(e)
        
        filename = f"chunk_{num_of_file + 7}_OLD_KATHARINA_{last_start}_end.json"
        with open(os.path.join(OUTPUT_DIR, filename), "w", encoding="utf-8") as out_f:
            json.dump(normalized_chunk_data, out_f, ensure_ascii=False, indent=2)
    
    print(f"Разделение завершено. Файлы сохранены в папке '{OUTPUT_DIR}/'")
    print(f'Общее количество outside меншенов полностью вне чанка: {DOC_COMP_OUTSIDE_M}')
    print(f'Общее количество outside меншенов частично вне чанка: {DOC_PARTLY_OUTSIDE_M}')    


            
        


# def main():
    
#     with open("Das Erdbeben in Chili.json", "r", encoding="utf-8") as f:
#         data = json.load(f)
        
#     mentions = data["mentions"]
#     participations = data["participations"]
    
#     full_annotations = combining_annotations_by_authors(data)
    
#     spans_with_tags = get_spans_with_tags(full_annotations)

    
#     insertions = get_insertions(spans_with_tags, full_annotations)
    
#     final_text = insert_tags_with_shift_and_save(full_annotations["text"], insertions)
    
#     new_json = format_text(full_annotations)
    
#     split_the_file()
            

# main()

INPUT_FOLDER = "input_texts"

def process_file(filepath, num_of_file):
    with open(filepath, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    mentions = data["mentions"]
    participations = data["participations"]
    
    full_annotations = combining_annotations_by_authors(data)
    
    #spans_with_tags = get_spans_with_tags(full_annotations)
    
    #insertions = get_insertions(spans_with_tags, full_annotations)
    
    #final_text = insert_tags_with_shift_and_save(full_annotations["text"], insertions, num_of_file)
    
    new_json = format_text(full_annotations, num_of_file)
    
    split_the_file(num_of_file)




def deviding_files_into_train_dev_test(BASE_DIR  = 'chunks', test_dev_only = True):
    TRAIN_DIR = os.path.join(BASE_DIR, "train")
    DEV_DIR   = os.path.join(BASE_DIR, "dev")
    TEST_DIR  = os.path.join(BASE_DIR, "test")
    
    # Создаём подпапки, если не существуют
    for folder in [TRAIN_DIR, DEV_DIR, TEST_DIR]:
        os.makedirs(folder, exist_ok=True)
    
    # for folder in [DEV_DIR, TEST_DIR]:
    #     os.makedirs(folder, exist_ok=True)
    
    # Получаем список всех файлов в chunks (исключая подкаталоги)
    all_files = [f for f in os.listdir(BASE_DIR) if os.path.isfile(os.path.join(BASE_DIR, f))]
    
    # Перемешиваем с фиксированным seed для воспроизводимости
    random.seed(42)
    random.shuffle(all_files)
    
    if test_dev_only:
        train_files = []
        dev_files, test_files = train_test_split(all_files, test_size=0.5, random_state=42)
        
    else:  
        # Сначала делим на train (80%) и temp (20%)
        train_files, temp_files = train_test_split(all_files, test_size=0.2, random_state=42)
        
        # Потом temp делим поровну на dev и test (10% + 10%)
        dev_files, test_files = train_test_split(temp_files, test_size=0.5, random_state=42)
    
    # Функция перемещения файлов
    def move_files(file_list, target_dir):
        for filename in file_list:
            shutil.move(os.path.join(BASE_DIR, filename), os.path.join(target_dir, filename))
    
    # Перемещаем
    move_files(train_files, TRAIN_DIR)
    move_files(dev_files, DEV_DIR)
    move_files(test_files, TEST_DIR)
    
    # Проверка
    total = len(all_files)
    print(f"Файлов всего:     {total}")
    print(f"→ Тренировка:     {len(train_files)}")
    print(f"→ Валидация (dev):{len(dev_files)}")
    print(f"→ Тест:           {len(test_files)}")
    print(f"Всего после разбиения: {len(train_files) + len(dev_files) + len(test_files)}")


#getting count of empty phrases (without any participants in agentive, low_agentive and passive)
def get_phrases_with_no_participants(formated_full_annotations):
    count = 0 
    for elem in formated_full_annotations['annotations']:
        if all(not elem[field]["spans"] for field in ("agentive", "low_agentive", "passive")):
            count += 1
            
    return count




def calc_avg_tokens_per_sent():
    formated_texts_dir = 'formated_full_texts'
    TOTAL_TOKENS = 0
    TOTAL_SENTENCES = 0
    tok = MT5Tokenizer.from_pretrained("google/mt5-xl")
    for file_name in os.listdir(formated_texts_dir):
        file_path = os.path.join(formated_texts_dir, file_name)
        with open(file_path, 'r', encoding='UTF-8') as f:
            data = json.load(f)
            sentences = data["sentences"]
            print(f'{data["documentName"]}: {len(sentences)} sentences')
            TOTAL_SENTENCES += len(sentences)
            #++++++считаем токены для предложений++++++++
            for sent in sentences:  # твой список предложений
                #lengths.append(len(tok.encode(sent["text"], add_special_tokens=False)))
                token_ids = tok.encode(sent["text"], add_special_tokens=False)
                TOTAL_TOKENS += len(token_ids)
                # print(token_ids)
                # print(tok.convert_ids_to_tokens(token_ids))
    AVG_TOKENS_PER_SENT = TOTAL_TOKENS / TOTAL_SENTENCES
    print(f'Avg tokens per sentences: {AVG_TOKENS_PER_SENT}')
    

def check_dev_test_chunks():
    file_path = 'dev-predicts.jsonlines'
    TOTAL_COUNT = 0 #total count of missed phrases in predictions, located after last predicted phrase span
    with open(file_path, 'r', encoding='UTF-8') as f:
        for line in f:
            data = json.loads(line)
            gold_clusters = data["gold_clusters"]
            predicted_clusters = data["predict_clusters"]
            cur_count = 0
            
            if not predicted_clusters:
                print(f'\nChunk {data["doc_key"]}: the clusters are empty!')
                continue
            
            predicted_phrase_triplets = [
                triplet 
                for cluster in predicted_clusters
                for triplet in cluster 
                if triplet[2] == 1
            ]
            
            gold_phrase_triplets = [
                triplet 
                for cluster in gold_clusters
                for triplet in cluster 
                if triplet[2] == 1
            ]
            
            last_predicted_phrase_begin = predicted_phrase_triplets[-1][0]
            last_predicted_phrase_end = predicted_phrase_triplets[-1][1]
            print(f'\nCHUNK: {data["doc_key"]}')
            print(f'Last predicted phrase: [{last_predicted_phrase_begin}, {last_predicted_phrase_end}]')
            
            for triplet in gold_phrase_triplets:
                if triplet[0] > last_predicted_phrase_end:
                    cur_count += 1 
                    print(f'gold phrase missed: [{triplet[0]}, {triplet[1]}]')
            
            TOTAL_COUNT += cur_count
    print(f'Total gold phrases missed after last phrase was predicted: {TOTAL_COUNT}')
            
            
        
            
    

def main():
    delete_old_chunks() # удаляем старые файла вывода
    num_of_file = 1
    for filename in os.listdir(INPUT_FOLDER):
        if filename.endswith(".json"):
            filepath = os.path.join(INPUT_FOLDER, filename)
            print(f"🔄 Обработка файла: {filename}")
            process_file(filepath, num_of_file)
            num_of_file += 1
    
    #deviding_files_into_train_dev_test()
        
    #calc_avg_tokens_per_sent()
    
    #check_dev_test_chunks()
    

main()











