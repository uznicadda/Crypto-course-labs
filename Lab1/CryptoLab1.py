import math
import re
from collections import defaultdict
import pandas as pd
import openpyxl

def process_text(text):
    text = text.lower()
    text = re.sub(r'[^а-яё\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    processed_text = text.lower()
    processed_text = processed_text.replace('ё', 'е').replace('ъ', 'ь')
    processed_text = ''.join(c for c in processed_text if c.isalpha() or c.isspace())
    processed_text = ' '.join(processed_text.split())
    return processed_text

def calculate_entropy(frequencies, total_symbols):
    entropy = 0
    for freq in frequencies.values():
        if freq > 0:
            p_i = freq / total_symbols
            entropy -= p_i * math.log2(p_i)
    return entropy

def create_and_print_bigram_matrix(bigram_frequencies, alphabet, filepath):
    matrix_data = defaultdict(lambda: defaultdict(float))
    total_bigrams = sum(bigram_frequencies.values())

    for bigram, freq in bigram_frequencies.items():
        first_char = bigram[0]
        second_char = bigram[1]
        matrix_data[first_char][second_char] = freq / total_bigrams

    df = pd.DataFrame(matrix_data).fillna(0)

    alphabet_list = list(alphabet)
    df = df.reindex(index=alphabet_list, columns=alphabet_list, fill_value=0)

    df.to_excel(filepath, index=True)

def create_and_print_unigram_table(unigram_frequencies, total_symbols, filepath):
    df = pd.DataFrame({
        'Буква': list(unigram_frequencies.keys()),
        'Кількість': list(unigram_frequencies.values()),
        'Частота': [freq / total_symbols for freq in unigram_frequencies.values()]
    })
    df = df.sort_values(by='Частота', ascending=False).reset_index(drop=True)

    df.to_excel(filepath, index=False)

def calculate_redundancy(H_inf, H0):
    return 1 - (H_inf / H0)

def main():
    filename = "./graf-monte-kristo.txt"
    try:
        with open(filename, "r", encoding="utf-8") as f:
            raw_text = f.read()
    except FileNotFoundError:
        print(f"Ошибка: файл по пути '{filename}' не найден.")
        return
    except UnicodeDecodeError:
        print("Ошибка кодировки: попробуйте изменить 'utf-8' на 'cp1251' или 'cp866'.")
        return

    processed_text_with_spaces = process_text(raw_text)
    processed_text_no_spaces = processed_text_with_spaces.replace(" ", "")

    with open("cleaned_text.txt", "w", encoding="utf-8") as f:
        f.write(processed_text_with_spaces)

    unigram_freq_with_spaces = defaultdict(int)
    for c in processed_text_with_spaces:
        unigram_freq_with_spaces[c] += 1
    total_symbols_with_spaces = len(processed_text_with_spaces)
    h1_with_spaces = calculate_entropy(unigram_freq_with_spaces, total_symbols_with_spaces)

    unigram_freq_no_spaces = defaultdict(int)
    for c in processed_text_no_spaces:
        unigram_freq_no_spaces[c] += 1
    total_symbols_no_spaces = len(processed_text_no_spaces)
    h1_no_spaces = calculate_entropy(unigram_freq_no_spaces, total_symbols_no_spaces)

    bigram_freq_with_spaces = defaultdict(int)
    for i in range(len(processed_text_with_spaces) - 1):
        bigram = processed_text_with_spaces[i:i+2]
        bigram_freq_with_spaces[bigram] += 1
    total_bigrams_with_spaces = len(processed_text_with_spaces) - 1
    h2_with_spaces = calculate_entropy(bigram_freq_with_spaces, total_bigrams_with_spaces) / 2

    bigram_freq_no_spaces = defaultdict(int)
    for i in range(len(processed_text_no_spaces) - 1):
        bigram = processed_text_no_spaces[i:i+2]
        bigram_freq_no_spaces[bigram] += 1
    total_bigrams_no_spaces = len(processed_text_no_spaces) - 1
    h2_no_spaces = calculate_entropy(bigram_freq_no_spaces, total_bigrams_no_spaces) / 2

    bigram_freq_with_spaces_step2 = defaultdict(int)
    for i in range(0, len(processed_text_with_spaces) - 1, 2):
        bigram = processed_text_with_spaces[i:i+2]
        bigram_freq_with_spaces_step2[bigram] += 1
    h2_step2_with_spaces = calculate_entropy(bigram_freq_with_spaces_step2, len(processed_text_with_spaces) // 2) / 2

    bigram_freq_no_spaces_step2 = defaultdict(int)
    for i in range(0, len(processed_text_no_spaces) - 1, 2):
        bigram = processed_text_no_spaces[i:i+2]
        bigram_freq_no_spaces_step2[bigram] += 1
    h2_step2_no_spaces = calculate_entropy(bigram_freq_no_spaces_step2, len(processed_text_no_spaces) // 2) / 2

    alphabet_with_space = 'абвгдежзийклмнопрстуфхцчшщыьэюя '
    alphabet_no_space = 'абвгдежзийклмнопрстуфхцчшщыьэюя'

    h0_with_spaces = math.log2(len(alphabet_with_space))
    h0_no_spaces = math.log2(len(alphabet_no_space))

    r_h1_with_spaces = calculate_redundancy(h1_with_spaces, h0_with_spaces)
    r_h1_no_spaces = calculate_redundancy(h1_no_spaces, h0_no_spaces)
    r_h2_with_spaces = calculate_redundancy(h2_with_spaces, h0_with_spaces)
    r_h2_no_spaces = calculate_redundancy(h2_no_spaces, h0_no_spaces)
    r_h2_step2_with_spaces = calculate_redundancy(h2_step2_with_spaces, h0_with_spaces)
    r_h2_step2_no_spaces = calculate_redundancy(h2_step2_no_spaces, h0_no_spaces)

    print("\n--- Значення ентропії ---")
    print(f"H1 (з пробілами): {h1_with_spaces:.5f}")
    print(f"H1 (без пробілів): {h1_no_spaces:.5f}")
    print(f"H2 (з пробілами): {h2_with_spaces:.5f}")
    print(f"H2 (без пробілів): {h2_no_spaces:.5f}")
    print(f"H2 з кроком 2 (з пробілами): {h2_step2_with_spaces:.5f}")
    print(f"H2 з кроком 2 (без пробілів): {h2_step2_no_spaces:.5f}")

    print("\n--- Значення надлишковості ---")
    print(f"R для H1 (з пробілами): {r_h1_with_spaces:.5f}")
    print(f"R для H1 (без пробілів): {r_h1_no_spaces:.5f}")
    print(f"R для H2 (з пробілами): {r_h2_with_spaces:.5f}")
    print(f"R для H2 (без пробілів): {r_h2_no_spaces:.5f}")
    print(f"R для H2 з кроком 2 (з пробілами): {r_h2_step2_with_spaces:.5f}")
    print(f"R для H2 з кроком 2 (без пробілів): {r_h2_step2_no_spaces:.5f}")

    create_and_print_unigram_table(unigram_freq_with_spaces, total_symbols_with_spaces, "./Unigram.xlsx")
    create_and_print_unigram_table(unigram_freq_no_spaces, total_symbols_no_spaces, "./UnigramNoSpace.xlsx")

    create_and_print_bigram_matrix(bigram_freq_with_spaces, alphabet_with_space, "./Bigram.xlsx")
    create_and_print_bigram_matrix(bigram_freq_with_spaces_step2, alphabet_with_space, "./BigramStep2.xlsx")
    create_and_print_bigram_matrix(bigram_freq_no_spaces, alphabet_with_space, "./BigramNoSpace.xlsx")
    create_and_print_bigram_matrix(bigram_freq_no_spaces_step2, alphabet_with_space, "./BigramStep2NoSpace.xlsx")


if __name__ == "__main__":
    main()