import pdfplumber
import re
import pandas as pd

roman_numerals = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
OPTION_LETTERS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
'''
 SECTION I INTRODUCTION TO CLINICAL MEDICINE
 Questions 1
 Answers 18
 SECTION II NUTRITION
 Questions 47
 Answers 50
 SECTION III ONCOLOGY AND HEMATOLOGY
 Questions 55
 Answers 71
 SECTION IV INFECTIOUS DISEASES
 Questions 103
 Answers 130
 SECTION V DISORDERS OF THE CARDIOVASCULAR SYSTEM
 Questions 175
 Answers 202
 SECTION VI DISORDERS OF THE RESPIRATORY SYSTEM
 Questions 237
 Answers 254
 SECTION VII DISORDERS OF THE URINARY AND KIDNEY TRACT
 Questions 283
 Answers 293
 SECTION VIII DISORDERS OF THE GASTROINTESTINAL SYSTEM
 Questions 307
 Answers 321
 SECTION IX RHEUMATOLOGY AND IMMUNOLOGY
 Questions 345
 Answers 358
 For more information about this title, click here
vi CONTENTS
 SECTION X ENDOCRINOLOGY AND METABOLISM
 Questions 379
 Answers 393
 SECTION XI NEUROLOGIC DISORDERS
 Questions 421
 Answers 435
 SECTION XII DERMATOLOGY
 Questions 457
 Answers 460

'''
chap = {'I': [18, 47],
        'II': [50, 55],
        'III': [71, 103],
        'IV': [130, 175],
        'V': [202, 237],
        'VI': [254, 283],
        'VII': [293, 307],
        'VIII': [321, 345],
        'IX': [358, 379],
        'X': [393, 421],
        'XI': [435, 457],
        'XII': [460, 465],
}

def is_option_start(line):
    return re.match(r'^[A-Ha-h]\.', line.strip())

def extract_options(lines, start_idx):
    options = {}
    current_option = None
    idx = start_idx

    while idx < len(lines):
        line = lines[idx].strip()

        # Stop if a new question begins
        if re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\.', line):
            break

        # Skip footer or copyright/junk
        if re.search(r'Copyright|Click here for terms|McGra|Hill|^\d+[\s\-]', line, re.IGNORECASE):
            idx += 1
            continue

        # Skip junk lines (e.g., CLINICAL MEDICINE)
        if line.isupper() and len(line.split()) <= 4:
            idx += 1
            continue

        # Start of new option
        match = re.match(r'^([A-Ha-h])\. +(.+)', line)
        if match:
            current_option = match.group(1).upper()
            options[current_option] = match.group(2).strip()
        elif current_option:
            # Merge hyphenated word across lines
            if options[current_option].endswith('-') and line and line[0].islower():
                options[current_option] = options[current_option][:-1] + line.strip()
            else:
                options[current_option] += ' ' + line.strip()

        idx += 1

    return options, idx

def clean_question_text(text):
    print("\n--- Raw Collected Text ---")
    print(text)

    

    # 2. Remove known footers and repeated irrelevant phrases line by line
    lines = text.split('\n')
    # print("\n--- Lines ---", lines)
    cleaned_lines = []
    for line in lines:
        if re.search(r'CLINICAL MEDICINE|Hill Companies', line, re.IGNORECASE):
            print(f"❌ Skipping footer line: {line}")
            continue

        cleaned_lines.append(line.strip())

    text = '. '.join(cleaned_lines)
    text = re.sub(r'Copyright ©.*?Click here for terms of use\.?', '', text, flags=re.IGNORECASE)
    print("\n--- After Footer Removal ---")
    print(text)

    # 3. Remove (continued)
    text = re.sub(r'\(\s*continued\s*\)', '', text, flags=re.IGNORECASE)
    print("\n--- After Continued Removal ---")
    print(text)

    # 1. Fix hyphenated line breaks FIRST
    text = text.replace('-\n', '').replace('\n', ' ')
    text = re.sub(r'(\w)-\s+(\w)', r'\1\2', text)  # Also fix leftover "- word" cases
    print("\n--- After Hyphen Fix ---")
    print(text)

    return text.strip()


def extract_questions_columnwise(pdf_path, start_page, end_page):
    questions = {}
    question_number = None
    current_question_lines = []
    is_question_section = True

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[start_page:end_page]:  # Adjust page range as needed
            print(f"\n📄 Page {page.page_number} Text:")
            
            header = 50
            footer = 0
            full_text = page.within_bbox((0, header, page.width, page.height - footer)).extract_text() or ''
            all_lines = full_text.split('\n')

            i = 0
            while i < len(all_lines):
                line = all_lines[i].strip()
                if page.page_number == 26:
                    print("Line", i, ":", line)

                if not is_question_section:
                    i += 1
                    continue

                # if re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\. The answer is [A-Ha-h]\.', line) or re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d{1,3}\. and (?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d{1,3}\. The answers are [A-H] and [A-H]\.', line):
                #     print(f"❌ Skipping answer line: {line}")
                #     i += 1
                #     continue


                # If an options block begins
                if is_option_start(line):
                    options_dict, new_idx = extract_options(all_lines, i)
                    print(f"📌 Collected options block: {options_dict}")
                    i = new_idx
                    continue

                # Match question start
                match = re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-(\d+)\.', line)
                continued_match = re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-(\d+)\. \(Continued\)', line, re.IGNORECASE)

                if continued_match:
                    cont_qnum = f"{continued_match.group(1)}-{continued_match.group(2)}"
                    if question_number == cont_qnum:
                        cleaned_line = re.sub(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\. \(Continued\)', '', line, flags=re.IGNORECASE).strip()
                        current_question_lines.append(cleaned_line)
                        print(f"🔁 Appending to continued: {cont_qnum}")
                    i += 1
                    continue

                if match:
                    new_qnum = f"{match.group(1)}-{match.group(2)}"

                    if new_qnum == question_number:
                        cleaned_line = re.sub(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\.', '', line).strip()
                        current_question_lines.append(cleaned_line)
                        print(f"🔁 Duplicate header in same question: {new_qnum}")
                        i += 1
                        continue

                    # Save previous question
                    if question_number and current_question_lines:
                        full_text = ' '.join(current_question_lines)
                        words = full_text.split()

                        if len(words) > 4:
                            first_part = ' '.join(words[:4])
                            rest_part = ' '.join(words[4:])
                            question_label_pattern = re.escape(question_number) + r'\.\s*'
                            rest_part = re.sub(question_label_pattern, '', rest_part)
                            full_text = f"{first_part} {rest_part}"
                        else:
                            full_text = ' '.join(words)

                        cleaned = clean_question_text(full_text)
                        print(f"📝 Cleaning question text: {cleaned}")
                        if question_number in questions:
                            questions[question_number] += ' ' + cleaned
                            print(f"🔄 Appended to existing question: {question_number}")
                        else:
                            questions[question_number] = cleaned
                            print(f"✅ New question added: {question_number}")

                    question_number = new_qnum
                    current_question_lines = [line]
                    print(f"▶️ Starting new question: {question_number}")
                elif question_number:
                    dup_header_inside = re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\.', line)
                    if dup_header_inside:
                        line = re.sub(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)-\d+\.', '', line).strip()
                        print(f"🧹 Removed duplicate header in body of {question_number}")
                    current_question_lines.append(line)

                i += 1

    # Save last question
    if question_number and current_question_lines:
        full_text = ' '.join(current_question_lines)
        words = full_text.split()

        if len(words) > 4:
            first_part = ' '.join(words[:4])
            rest_part = ' '.join(words[4:])
            question_label_pattern = re.escape(question_number) + r'\.\s*'
            rest_part = re.sub(question_label_pattern, '', rest_part)
            full_text = f"{first_part} {rest_part}"
        else:
            full_text = ' '.join(words)

        cleaned = clean_question_text(full_text)
        print(f"📝 Cleaning question text 1: {cleaned}")
        if question_number in questions:
            questions[question_number] += ' ' + cleaned
            print(f"🔄 Appended to existing question: {question_number}")
        else:
            questions[question_number] = cleaned
            print(f"✅ New question added: {question_number}")

    print(f"\n📦 Total questions collected: {len(questions)}")
    return [questions[q] for q in sorted(
        questions.keys(),
        key=lambda x: (roman_numerals.index(x.split('-')[0]), int(x.split('-')[1]))
    )]

def save_to_excel(questions, excel_path):
    df = pd.DataFrame(questions, columns=["Questions"])
    df.to_excel(excel_path, index=False)
    print(f"\n📥 Saved {len(questions)} questions to {excel_path}")
    print(df.head())

if __name__ == "__main__":
    pdf_file = r"C:\Users\kavya\Downloads\Harrison Self Assessment, 17th.pdf"
    dfs = {}
    for roman_number, chapter in chap.items():
        start_page, end_page = chapter
        print(f"Processing {roman_number}: {chapter}")
        excel_file = f"questions_{roman_number}.xlsx"
        questions = extract_questions_columnwise(pdf_file, start_page + 10, end_page + 10)
        df = pd.DataFrame(questions, columns=["Questions"])
        dfs[roman_number] = df
        save_to_excel(questions, excel_file)

    for rn, df in dfs.items():
        print(f"\n📊 DataFrame for {rn}:")
        print(df.tail())
        print(f"Total questions: {len(df)}")

    # Concatenate all DataFrames (outside the loop!)
    final_df = pd.concat(dfs.values(), ignore_index=True)
    excel_file = "final_questions.xlsx"
    save_to_excel(final_df, excel_file)
    print(f"\n📥 Saved final questions to {excel_file}")

