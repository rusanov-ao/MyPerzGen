from docx import Document

def parse_word_document(file_path):
    document = Document(file_path)
    slides = []

    current_slide = None

    for paragraph in document.paragraphs:
        # Пропускаем пустые параграфы
        if not paragraph.text.strip():
            continue

        # Проверяем, если параграф является заголовком
        if paragraph.style.name.startswith('Heading'):
            if current_slide:
                slides.append(current_slide)
            current_slide = {
                'title': paragraph.text,
                'paragraphs': [],
                'text_as_is': []
            }
        elif current_slide:
            # Проверяем, если параграф выделен курсивом
            is_italic = any(run.italic for run in paragraph.runs)
            if is_italic:
                current_slide['text_as_is'].append(paragraph.text)
            else:
                current_slide['paragraphs'].append(paragraph.text)

    # Добавляем последний слайд в список
    if current_slide:
        slides.append(current_slide)

    return slides

# Пример использования
if __name__ == "__main__":
    file_path = "C:/Users/Администратор/Desktop/MyPerzGen/MyPic/document.docx"
    slides = parse_word_document(file_path)

    # Выводим информацию о слайдах
    for slide in slides:
        print(f"Заголовок: {slide['title']}")
        print("Параграфы для суммаризации:")
        for para in slide['paragraphs']:
            print(f"Параграф: {para}")
        print("Текст как есть:")
        for para in slide['text_as_is']:
            print(f"Параграф: {para}")
        print("\n" + "-"*50 + "\n")
