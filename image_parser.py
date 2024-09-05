import os


def scan_image_folder(folder_path):
    images = {}

    # Получаем список всех файлов в папке
    for filename in os.listdir(folder_path):
        # Проверяем, является ли файл изображением и подходит ли под шаблон
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            # Извлекаем номер из названия файла
            try:
                slide_number = int(os.path.splitext(filename)[0])
                images[slide_number] = filename
            except ValueError:
                # Пропускаем файлы, которые не содержат числового названия
                continue

    return images


# Пример использования
if __name__ == "__main__":
    folder_path = "C:/Users/Администратор/Desktop/MyPerzGen/MyPic"
    images = scan_image_folder(folder_path)

    for slide_number, image_name in images.items():
        print(f"Слайд {slide_number}: {image_name}")
