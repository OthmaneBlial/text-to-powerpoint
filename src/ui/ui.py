from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QTextEdit, QPushButton, QHBoxLayout, 
    QFileDialog, QMessageBox, QComboBox, QFontComboBox, QSpinBox, QColorDialog, 
    QCheckBox, QTabWidget
)
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import Qt, QSize
from core.generator import SlideGenerator
from templates.templates import TEMPLATES
from utils.utils import convert_pptx_to_image, load_image
import logging
import os

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Advanced PowerPoint Slide Generator")
        self.setGeometry(100, 100, 1000, 700)
        self.setStyleSheet("""
            QWidget {
                background-color: #f9f9f9;
            }
            QLabel {
                font-size: 18px;
                color: #333333;
            }
            QTextEdit {
                font-size: 14px;
                border: 1px solid #cccccc;
                border-radius: 5px;
            }
            QPushButton {
                font-size: 14px;
                padding: 8px 16px;
                border-radius: 5px;
            }
            QPushButton#generate {
                background-color: #4CAF50;
                color: white;
            }
            QPushButton#generate:hover {
                background-color: #45a049;
            }
            QPushButton#clear {
                background-color: #f44336;
                color: white;
            }
            QPushButton#clear:hover {
                background-color: #da190b;
            }
            QComboBox, QFontComboBox, QSpinBox {
                font-size: 14px;
                padding: 5px;
            }
            QCheckBox {
                font-size: 14px;
            }
            QLabel#preview_label {
                border: 1px solid #cccccc;
                border-radius: 5px;
                padding: 10px;
                background-color: #ffffff;
                min-height: 300px;
                max-height: 600px;
                max-width: 800px;
                alignment: Qt.AlignCenter;
            }
        """)

        self.layout = QVBoxLayout(self)
        
        # Title Label
        self.title_label = QLabel("Advanced PowerPoint Slide Generator")
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("font-size: 24px; font-weight: bold;")
        self.layout.addWidget(self.title_label)
        
        # Tabs
        self.tabs = QTabWidget()
        self.layout.addWidget(self.tabs)
        
        # Input Tab
        self.input_tab = QWidget()
        self.tabs.addTab(self.input_tab, "Input")
        self.input_layout = QVBoxLayout(self.input_tab)
        
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("Enter your presentation content using Markdown syntax...\n\n"
                                         "# Slide Title\n"
                                         "## Slide Subtitle\n"
                                         "- Bullet Point 1\n"
                                         "- Bullet Point 2\n"
                                         "![Image Description](https://example.com/image.png)\n"
                                         "@chart {Categories: A,B,C; Series1: 10,20,30; Series2: 15,25,35}\n"
                                         "> Inspirational Quote")
        self.input_layout.addWidget(self.text_edit)
        
        # Template Tab
        self.template_tab = QWidget()
        self.tabs.addTab(self.template_tab, "Template")
        self.template_layout = QVBoxLayout(self.template_tab)
        
        # Template Selection
        self.template_selection_layout = QHBoxLayout()
        self.template_selection_layout.addWidget(QLabel("Select Template:"))
        self.template_combo = QComboBox()
        self.template_combo.addItems(TEMPLATES.keys())
        self.template_combo.currentTextChanged.connect(self.load_template_settings)
        self.template_selection_layout.addWidget(self.template_combo)
        self.template_layout.addLayout(self.template_selection_layout)
        
        # Customization Options
        self.customization_layout = QHBoxLayout()
        
        # Font Selection
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Calibri"))
        self.customization_layout.addWidget(QLabel("Font:"))
        self.customization_layout.addWidget(self.font_combo)
        
        # Title Font Size
        self.title_font_size_spin = QSpinBox()
        self.title_font_size_spin.setRange(10, 100)
        self.title_font_size_spin.setValue(40)
        self.customization_layout.addWidget(QLabel("Title Font Size:"))
        self.customization_layout.addWidget(self.title_font_size_spin)
        
        # Content Font Size
        self.content_font_size_spin = QSpinBox()
        self.content_font_size_spin.setRange(8, 72)
        self.content_font_size_spin.setValue(24)
        self.customization_layout.addWidget(QLabel("Content Font Size:"))
        self.customization_layout.addWidget(self.content_font_size_spin)
        
        self.template_layout.addLayout(self.customization_layout)
        
        # Color Pickers
        self.color_layout = QHBoxLayout()
        
        # Theme Color
        self.theme_color_btn = QPushButton("Select Theme Color")
        self.theme_color_btn.clicked.connect(self.select_theme_color)
        self.color_layout.addWidget(self.theme_color_btn)
        
        # Background Color
        self.bg_color_btn = QPushButton("Select Background Color")
        self.bg_color_btn.clicked.connect(self.select_background_color)
        self.color_layout.addWidget(self.bg_color_btn)
        
        self.template_layout.addLayout(self.color_layout)
        
        # Bold Options
        self.bold_layout = QHBoxLayout()
        self.title_bold_checkbox = QCheckBox("Bold Titles")
        self.title_bold_checkbox.setChecked(True)
        self.content_bold_checkbox = QCheckBox("Bold Content")
        self.bold_layout.addWidget(self.title_bold_checkbox)
        self.bold_layout.addWidget(self.content_bold_checkbox)
        self.template_layout.addLayout(self.bold_layout)
        
        # Apply Customizations Button
        self.apply_custom_btn = QPushButton("Apply Customizations")
        self.apply_custom_btn.clicked.connect(self.apply_customizations)
        self.template_layout.addWidget(self.apply_custom_btn)
        
        # Preview Tab
        self.preview_tab = QWidget()
        self.tabs.addTab(self.preview_tab, "Preview")
        self.preview_layout = QVBoxLayout(self.preview_tab)
        
        self.preview_label = QLabel("Preview will be available after generating the presentation.")
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setObjectName("preview_label")
        self.preview_layout.addWidget(self.preview_label)
        
        # Buttons Layout
        self.buttons_layout = QHBoxLayout()
        
        self.clear_button = QPushButton("Clear Input")
        self.clear_button.setObjectName("clear")
        self.clear_button.clicked.connect(self.clear_input)
        self.buttons_layout.addWidget(self.clear_button)
        
        self.generate_button = QPushButton("Generate Presentation")
        self.generate_button.setObjectName("generate")
        self.generate_button.clicked.connect(self.generate_presentation)
        self.buttons_layout.addWidget(self.generate_button)
        
        self.layout.addLayout(self.buttons_layout)
        
        # Initialize Template Settings
        self.load_template_settings(self.template_combo.currentText())
    
    def load_template_settings(self, template_name):
        template = TEMPLATES.get(template_name)
        if not template:
            return
        self.font_combo.setCurrentFont(QFont(template.font_family))
        self.title_font_size_spin.setValue(template.title_font_size)
        self.content_font_size_spin.setValue(template.content_font_size)
        self.theme_color = template.theme_color
        self.bg_color = template.background_color
        self.update_color_buttons()
    
    def update_color_buttons(self):
        theme_hex = self.rgb_to_hex(self.theme_color)
        bg_hex = self.rgb_to_hex(self.bg_color)
        self.theme_color_btn.setStyleSheet(f"background-color: {theme_hex}")
        self.bg_color_btn.setStyleSheet(f"background-color: {bg_hex}")
    
    def rgb_to_hex(self, color):
        return '#{:02X}{:02X}{:02X}'.format(*color)
    
    def select_theme_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.theme_color = (color.red(), color.green(), color.blue())
            self.theme_color_btn.setStyleSheet(f"background-color: {color.name()}")
    
    def select_background_color(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.bg_color = (color.red(), color.green(), color.blue())
            self.bg_color_btn.setStyleSheet(f"background-color: {color.name()}")
    
    def apply_customizations(self):
        template_name = self.template_combo.currentText()
        template = TEMPLATES.get(template_name)
        if not template:
            QMessageBox.critical(self, "Error", "Selected template not found.")
            return
        template.font_family = self.font_combo.currentFont().family()
        template.title_font_size = self.title_font_size_spin.value()
        template.content_font_size = self.content_font_size_spin.value()
        template.theme_color = self.theme_color
        template.background_color = self.bg_color
        template.title_bold = self.title_bold_checkbox.isChecked()
        template.content_bold = self.content_bold_checkbox.isChecked()
        QMessageBox.information(self, "Success", "Template customizations applied successfully.")
    
    def clear_input(self):
        self.text_edit.clear()
        self.preview_label.setText("Preview will be available after generating the presentation.")
    
    def generate_presentation(self):
        input_text = self.text_edit.toPlainText()
        if not input_text.strip():
            QMessageBox.critical(self, "Error", "Please enter some text for the slides.")
            return

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Presentation", "", "PowerPoint Presentation (*.pptx)", options=options
        )
        if not file_path:
            return
        if not file_path.lower().endswith('.pptx'):
            file_path += '.pptx'

        template_name = self.template_combo.currentText()
        template = TEMPLATES.get(template_name)
        if not template:
            QMessageBox.critical(self, "Error", "Selected template not found.")
            return

        generator = SlideGenerator(template)
        try:
            generator.generate_presentation(input_text, file_path)
            QMessageBox.information(self, "Success", f"Presentation saved as {file_path}")
            self.preview_presentation(file_path)
        except Exception as e:
            logging.error(f"Failed to generate presentation: {e}", exc_info=True)
            QMessageBox.critical(self, "Error", f"An error occurred: {str(e)}")

    def preview_presentation(self, file_path):
        temp_image = os.path.join(os.path.dirname(file_path), "temp_preview.png")
        success = convert_pptx_to_image(file_path, temp_image)
        if success and os.path.exists(temp_image):
            pixmap = load_image(temp_image)
            if pixmap:
                self.preview_label.setPixmap(pixmap.scaled(
                    QSize(800, 600), Qt.KeepAspectRatio, Qt.SmoothTransformation))
                os.remove(temp_image)
                return
        self.preview_label.setText("Failed to generate preview image.")
