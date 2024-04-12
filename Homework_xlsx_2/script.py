import xlsxwriter

def get_content(file_name):
    with open(file_name) as file:
        return file.read()

def count_items(content):
    counts = {"Vowels": {}, "Consonants": {}, "Numbers": {}, "Symbols": {}}
    vowels = "aeiou"
    for char in content.lower():
        if char.isalpha():
            category = "Vowels" if char in vowels else "Consonants"
        elif char.isdigit():
            category = "Numbers"
        else:
            category = "Symbols"
        counts[category][char] = counts[category].get(char, 0) + 1
    return counts

def write_to_excel(file_name, counts):
    workbook = xlsxwriter.Workbook(file_name)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({"bold": True})
    vowel_format = workbook.add_format({"bg_color": "#FFFF00"})
    consonant_format = workbook.add_format({"bg_color": "#00FF00"})
    number_format = workbook.add_format({"bg_color": "#00FFFF"})
    symbol_format = workbook.add_format({"bg_color": "#FF00FF"})
    
    row = 0
    for category, items in counts.items():
        color_format = None
        if category == "Vowels":
            color_format = vowel_format
        elif category == "Consonants":
            color_format = consonant_format
        elif category == "Numbers":
            color_format = number_format
        else:
            color_format = symbol_format
            
        worksheet.write(row, 0, category, bold)
        col = 1
        for char, count in items.items():
            worksheet.write(row, col, f"{char} :      {count}", color_format)
            col += 1
        row += 1
    workbook.close()

def main():
    fname = "data.txt"
    output = "resources.xlsx"
    content = get_content(fname)
    counts = count_items(content)
    write_to_excel(output, counts)

if __name__ == "__main__":
    main()
