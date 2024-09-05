import time
from transformers import T5ForConditionalGeneration
from transformers import T5Tokenizer


def summarize_text(text, model_directory='/home/alex/PycharmProjects/pythonProject4/rut5_base_sum_gazeta', max_length=100, min_length=40):
    """
    Выполняет суммаризацию текста с использованием модели T5.

    :param text: Текст для суммаризации.
    :param model_directory: Путь к локальной директории с моделью и токенизатором.
    :param max_length: Максимальная длина суммаризированного текста.
    :param min_length: Минимальная длина суммаризированного текста.
    :return: Суммаризированный текст.
    """
    start_time = time.time()
    print("Загружается модель и токенизатор...")

    # Загрузка модели и токенизатора
    tokenizer = T5Tokenizer.from_pretrained(model_directory, legacy=True, clean_up_tokenization_spaces=True)
    model = T5ForConditionalGeneration.from_pretrained(model_directory)
    print(f"Модель и токенизатор загружены. Время: {time.time() - start_time:.2f} секунд.")

    # Подсчет количества токенов
    num_tokens = len(tokenizer.encode(text, truncation=False))
    print(f"Количество токенов в тексте: {num_tokens}")

    # Подготовка входных данных для суммаризации
    inputs = tokenizer.encode("summarize: " + text, return_tensors="pt", max_length=1024, truncation=True)

    # Генерация суммаризации
    print("Запуск генерации суммаризации...")
    summary_ids = model.generate(inputs, max_length=max_length, min_length=min_length, length_penalty=2.0,
                                 num_beams=4, early_stopping=True)
    summary = tokenizer.decode(summary_ids[0], skip_special_tokens=True)

    print(f"Суммаризация завершена. Время: {time.time() - start_time:.2f} секунд.")
    return summary


