import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QFileDialog, QVBoxLayout, QHBoxLayout, QWidget, QTextEdit
from summarizer import summarize_text
from word_parser import parse_word_document
from image_parser import scan_image_folder
from generate_pptx import save_presentation_from_data

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Выбор папок и документа')
        self.setGeometry(100, 100, 1200, 500)

        main_layout = QHBoxLayout()

        # Левая часть с кнопками и метками
        layout = QVBoxLayout()

        # Захардкоженные пути
        self.default_save_path = "/home/alex/PycharmProjects/pythonProject4/MyRez"
        self.default_images_path = "/home/alex/PycharmProjects/pythonProject4/MyPic"
        self.default_word_path = "/home/alex/PycharmProjects/pythonProject4/MyPic/document.docx"

        # Кнопка и метка для выбора папки сохранения презентации
        self.save_path_label = QLabel(f'Папка для сохранения: {self.default_save_path}', self)
        layout.addWidget(self.save_path_label)

        self.save_button = QPushButton('Выбрать папку для сохранения', self)
        self.save_button.clicked.connect(self.select_save_folder)
        layout.addWidget(self.save_button)

        # Кнопка и метка для выбора Word документа
        self.word_path_label = QLabel(f'Документ Word: {self.default_word_path}', self)
        layout.addWidget(self.word_path_label)

        self.word_button = QPushButton('Выбрать документ Word', self)
        self.word_button.clicked.connect(self.select_word_file)
        layout.addWidget(self.word_button)

        # Кнопка и метка для выбора папки с изображениями
        self.images_path_label = QLabel(f'Папка с изображениями: {self.default_images_path}', self)
        layout.addWidget(self.images_path_label)

        self.images_button = QPushButton('Выбрать папку с изображениями', self)
        self.images_button.clicked.connect(self.select_images_folder)
        layout.addWidget(self.images_button)

        # Кнопка для чтения текста
        self.readword_button = QPushButton('Прочитать текст', self)
        self.readword_button.clicked.connect(self.process_word_document)
        layout.addWidget(self.readword_button)

        # Кнопка для генерации презентации
        self.generate_button = QPushButton('Генерировать презентацию', self)
        self.generate_button.clicked.connect(self.generate_presentation)
        layout.addWidget(self.generate_button)

        # Кнопка для сохранения презентации в файл .pptx
        self.save_pptx_button = QPushButton('Сохранить презентацию в .pptx', self)
        self.save_pptx_button.clicked.connect(self.save_presentation)
        layout.addWidget(self.save_pptx_button)

        # Текстовое поле "Ход работы"
        self.progress_text = QTextEdit(self)
        self.progress_text.setReadOnly(True)
        layout.addWidget(QLabel("Ход работы:", self))
        layout.addWidget(self.progress_text)

        main_layout.addLayout(layout)

        # Новое текстовое поле для отображения исходного словаря
        source_data_layout = QVBoxLayout()
        source_data_layout.addWidget(QLabel("Исходные данные:", self))
        self.source_data_text = QTextEdit(self)
        self.source_data_text.setReadOnly(True)
        source_data_layout.addWidget(self.source_data_text)
        main_layout.addLayout(source_data_layout)

        # Текстовое поле для отображения результата
        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        main_layout.addWidget(self.output_text)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

        # Установка значений из захардкоженных путей
        self.save_path = self.default_save_path
        self.images_path = self.default_images_path
        self.word_path = self.default_word_path
        self.presentation_data = None

    def select_save_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выбрать папку для сохранения')
        if folder:
            self.save_path_label.setText(f'Папка для сохранения: {folder}')
            self.save_path = folder
            self.progress_text.append(f"Папка для сохранения выбрана: {folder}")

    def select_images_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Выбрать папку с изображениями')
        if folder:
            self.images_path_label.setText(f'Папка с изображениями: {folder}')
            self.images_path = folder
            self.progress_text.append(f"Папка с изображениями выбрана: {folder}")

    def select_word_file(self):
        file, _ = QFileDialog.getOpenFileName(self, 'Выбрать документ Word', '', 'Word Files (*.docx)')
        if file:
            self.word_path_label.setText(f'Документ Word: {file}')
            self.word_path = file
            self.progress_text.append(f"Документ Word выбран: {file}")

    def process_word_document(self):
        self.progress_text.append("Начинается обработка документа Word...")
        self.slides = parse_word_document(self.word_path)
        self.progress_text.append("Документ Word успешно обработан.")
        self.display_source_data()

    # def generate_presentation(self):
    #     self.progress_text.append("Начинается генерация презентации...")
    #
    #     images = scan_image_folder(self.images_path)
    #     self.progress_text.append("Изображения успешно загружены.")
    #
    #     self.presentation_data = []
    #     threshold_length = 524  # Минимальная длина текста для суммаризации
    #
    #     for i, slide in enumerate(self.slides):
    #         slide_number = i + 1
    #         summarized_paragraphs = []
    #
    #         for para in slide['paragraphs']:
    #             if len(para) >= threshold_length:
    #                 summarized_text = summarize_text(para)
    #                 summarized_paragraphs.append(summarized_text)
    #             else:
    #                 summarized_paragraphs.append(para)  # Если параграф слишком короткий, добавляем его без изменений
    #
    #         slide_data = {
    #             'title': slide['title'],
    #             'paragraphs': summarized_paragraphs,
    #             'text_as_is': slide.get('text_as_is', []),
    #             'image': images.get(slide_number, None)
    #         }
    #
    #         self.presentation_data.append(slide_data)
    #         self.progress_text.append(f"Слайд {slide_number} обработан.")
    #
    #     self.progress_text.append("Генерация презентации завершена.")

    def generate_presentation(self):
        self.progress_text.append("Начинается генерация презентации...")

        images = scan_image_folder(self.images_path)
        self.progress_text.append("Изображения успешно загружены.")

        self.presentation_data = []
        threshold_length = 300  # Минимальная длина текста для суммаризации

        for i, slide in enumerate(self.slides):
            slide_number = i + 1
            summarized_paragraphs = []

            for para in slide['paragraphs']:
                if len(para) >= threshold_length:
                    summarized_text = summarize_text(para)
                    summarized_paragraphs.append(summarized_text)
                else:
                    summarized_paragraphs.append(para)  # Если параграф слишком короткий, добавляем его без изменений

            slide_data = {
                'title': slide['title'],
                'paragraphs': summarized_paragraphs,  # Просто добавляем текст параграфов
                'text_as_is': slide.get('text_as_is', []),
                'image': images.get(slide_number, None)
            }

            self.presentation_data.append(slide_data)
            self.progress_text.append(f"Слайд {slide_number} обработан.")

            # Обновление окна с результатами
            self.output_text.append(f"Слайд: {slide_data['title']}")
            for para in slide_data['paragraphs']:
                self.output_text.append(f"    {para}")
            if slide_data['text_as_is']:
                for para in slide_data['text_as_is']:
                    self.output_text.append(f"    {para}")
            if slide_data['image']:
                self.output_text.append(f"  Изображение: {slide_data['image']}")
            else:
                self.output_text.append("  Изображение: нет")

            self.output_text.append("")  # Пустая строка между слайдами

            QApplication.processEvents()  # Обновление интерфейса

        self.progress_text.append("Генерация презентации завершена.")
        # Отображаем данные презентации в текстовом поле
        self.display_presentation_data()

    def display_source_data(self):
        self.source_data_text.clear()
        for slide in self.slides:
            self.source_data_text.append(f"Слайд: {slide['title']}")
            if slide['paragraphs']:
                self.source_data_text.append("  Параграфы для суммаризации:")
                for para in slide['paragraphs']:
                    self.source_data_text.append(f"    {para}")
            if slide.get('text_as_is', []):
                self.source_data_text.append("  Текст как есть:")
                for para in slide['text_as_is']:
                    self.source_data_text.append(f"    {para}")
            self.source_data_text.append("")

    def display_presentation_data(self):
        self.output_text.clear()
        for slide in self.presentation_data:
            self.output_text.append(f"Слайд: {slide['title']}")
            if slide['paragraphs']:
                self.output_text.append("  Параграфы для суммаризации:")
                for para in slide['paragraphs']:
                    self.output_text.append(f"    {para}")
            if slide['text_as_is']:
                self.output_text.append("  Текст как есть:")
                for para in slide['text_as_is']:
                    self.output_text.append(f"    {para}")
            if slide['image']:
                self.output_text.append(f"  Изображение: {slide['image']}")
            else:
                self.output_text.append("  Изображение: нет")
            self.output_text.append("")

    def save_presentation(self):
        if not self.presentation_data:
            self.progress_text.append("Ошибка: Нет данных для сохранения.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, 'Сохранить презентацию как', self.save_path, 'PowerPoint Files (*.pptx)')
        if file_path:
            save_presentation_from_data(self.presentation_data, file_path, self.images_path)
            self.progress_text.append(f"Презентация сохранена в {file_path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
