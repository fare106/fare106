from ast import main
import sys
import win32com.client
from PyQt5 import QtWidgets, uic, QtCore
from PyQt5.QtWidgets import QPushButton, QVBoxLayout
from PyQt5.QtGui import QIcon
import openpyxl
import os
import win32com.client


import sys

# from PyQt5 import QtWidgets"""  """


class MyApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(MyApp, self).__init__()
        self.setupUi(self)

    def setupUi(self, MainWindow):
        # Butonları tanımlama kodunuz

        def __init__(self):
            super().__init__()

        uic.loadUi("C:/pyt/qt/for.ui", self)

        # Başlangıçta genişlik ve yükseklik sıfır
        self.setGeometry(0, 0, 0, 950)  # Yüksekliği sabit tut
        self.shrink_width = 1150
        QtCore.QTimer.singleShot(0, self.start_expansion)

        # Butonları ve etiketleri bul
        self.button_open_1 = self.findChild(QtWidgets.QPushButton, "pushButton_1")
        self.button_open_2 = self.findChild(QtWidgets.QPushButton, "pushButton_2")
        self.button_open_3 = self.findChild(QtWidgets.QPushButton, "pushButton_3")
        self.button_open_4 = self.findChild(QtWidgets.QPushButton, "pushButton_4")
        self.button_open_5 = self.findChild(QtWidgets.QPushButton, "pushButton_5")
        self.button_open_6 = self.findChild(QtWidgets.QPushButton, "pushButton_A2")
        self.button_y1 = self.findChild(QtWidgets.QPushButton, "pushButton_Y1")
        self.button_y2 = self.findChild(QtWidgets.QPushButton, "pushButton_Y2")
        self.button_open_8 = self.findChild(QtWidgets.QPushButton, "pushButton_8")
        self.button_open_10 = self.findChild(QtWidgets.QPushButton, "pushButton_10")
        self.button_open_11 = self.findChild(QtWidgets.QPushButton, "pushButton_11")
        self.button_open_12 = self.findChild(QtWidgets.QPushButton, "pushButton_12")

        # Etiketleri bul ve başlangıçta gizle
        self.label_1 = self.findChild(QtWidgets.QLabel, "label_1")
        self.label_2 = self.findChild(QtWidgets.QLabel, "label_2")
        self.label_3 = self.findChild(QtWidgets.QLabel, "label_3")
        self.label_4 = self.findChild(QtWidgets.QLabel, "label_4")
        self.label_5 = self.findChild(QtWidgets.QLabel, "label_5")
        self.label_6 = self.findChild(QtWidgets.QLabel, "label_6")
        self.label_10 = self.findChild(QtWidgets.QLabel, "label_10")
        self.label_11 = self.findChild(QtWidgets.QLabel, "label_11")
        self.label_12 = self.findChild(QtWidgets.QLabel, "label_12")

        # Yeni butonları oluştur ve yerleştir
        layout = QVBoxLayout()

        # Etiketlerin varlığını kontrol et
        labels = [
            self.label_1,
            self.label_2,
            self.label_3,
            self.label_4,
            self.label_5,
            self.label_6,
            self.label_10,
            self.label_11,
            self.label_12,
        ]
        for i, label in enumerate(labels, 1):
            if label is None:
                print(f"label_{i} bulunamadı. UI dosyasını kontrol edin.")
            else:
                label.setVisible(False)

        # Butonlar için tıklama olayları ekle

        if Self.button_open_1:

            self.button_open_1.clicked.connect(self.close)
            self.button_open_1.clicked.connect(
                lambda: self.open_excel("C:/İŞ DOSYASI/SETUP EXE/MAAS.xlsm")
            )

        if self.button_open_2:

            self.button_open_2.clicked.connect(self.close)
            self.button_open_2.clicked.connect(
                lambda: self.open_excel("C:/İŞ DOSYASI/SETUP EXE/KUMASBYFABRIC.xlsm")
            )

        if self.button_open_6:
            self.button_open_6.clicked.connect(self.close)
            self.button_open_6.clicked.connect(
                lambda: self.open_excel("C:/İŞ DOSYASI/SETUP EXE/MUSTERILER.xlsm")
            )
            self.button_open_6.clicked.connect(
                lambda: self.open_ui("C:/pyt/qt/gırıs.ui")
            )

        def open_ui(self, ui_path):
            # Burada UI dosyasını açmak için gerekli kodu ekleyin
            print(f"UI dosyası açılıyor: {ui_path}")

        if self.button_open_4:
            self.button_open_4.clicked.connect(self.close)
            self.button_open_4.clicked.connect(
                lambda: self.open_excel("C:/İŞ DOSYASI/SETUP EXE/kasa.xlsm")
            )
        # self.button_open_4.clicked.connect (uic.loadUi('C:/pyt/qt/as.ui', self))
        if self.button_open_5:
            self.button_open_5.clicked.connect(self.close)

            self.button_open_5.clicked.connect(
                lambda: self.open_excel("C:/İŞ DOSYASI/SETUP EXE/STOK.xlsm")
            )

        if self.button_y1:
            self.button_y1.clicked.connect(self.expand_window)
        if self.button_y2:
            self.button_y2.clicked.connect(self.shrink_window)

        if self.button_open_8:
            self.button_open_8.clicked.connect(self.close)

        # Butonlar için eventFilter'ları ekle
        for button in [
            self.button_open_1,
            self.button_open_2,
            self.button_open_3,
            self.button_open_4,
            self.button_open_5,
            self.button_open_6,
            self.button_open_10,
            self.button_open_11,
            self.button_open_12,
        ]:
            if button:
                button.installEventFilter(self)

        # Timer oluştur
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.update_size)
        self.is_expanding = False
        # Hedef boyutlar

        self.expand_width = 1550
        self.shrink_width = 1150
        self.fixed_height = 950  # Yükseklik sabit

        # Excel uygulamasını başlat
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = False  # Excel uygulamasını görünmez yap

        # İki değişken ekle
        self.is_expanding = False
        self.is_shrinking = False

        # Başlık çubuğunu kaldır
        self.setWindowFlags(QtCore.Qt.FramelessWindowHint)

    def start_expansion(self):
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        x_center = (screen.width() - self.expand_width) // 2
        y_center = (screen.height() - self.fixed_height) // 2
        self.setGeometry(x_center, y_center, 0, self.fixed_height)
        self.is_expanding = False
        self.timer.start(10)

    def eventFilter(self, source, event):
        if event.type() == QtCore.QEvent.Enter:
            if source == self.button_open_1:
                self.label_1.setVisible(True)
            elif source == self.button_open_2:
                self.label_2.setVisible(True)
            elif source == self.button_open_3:
                self.label_3.setVisible(True)
                self.is_expanding = True
                self.timer.start(10)
            elif source == self.button_open_4:
                self.label_4.setVisible(True)
                self.is_shrinking = True

            elif source == self.button_open_5:

                self.label_5.setVisible(True)
            elif source == self.button_open_6:
                self.label_6.setVisible(True)
            elif source == self.button_open_10:
                self.label_10.setVisible(True)

            elif source == self.button_open_11:
                self.label_11.setVisible(True)

            elif source == self.button_open_12:
                self.label_12.setVisible(True)

        elif event.type() == QtCore.QEvent.Leave:
            if source == self.button_open_1:
                self.label_1.setVisible(False)
            elif source == self.button_open_2:
                self.label_2.setVisible(False)
            elif source == self.button_open_3:
                self.label_3.setVisible(False)

                if not self.is_expanding:
                    self.is_shrinking = True
                    self.timer.start(10)
            elif source == self.button_open_4:
                self.label_4.setVisible(False)
                self.is_shrinking = True
                self.timer.start(10)
            elif source == self.button_open_5:
                self.label_5.setVisible(False)
            elif source == self.button_open_6:
                self.label_6.setVisible(False)

            elif source == self.button_open_10:
                self.label_10.setVisible(False)

            elif source == self.button_open_11:
                self.label_11.setVisible(False)
            elif source == self.button_open_12:
                self.label_12.setVisible(False)

        return super().eventFilter(source, event)

    def update_size(self):
        current_width = self.width()
        step = 10  # Adım büyüklüğü

        if self.is_expanding:
            if current_width < self.expand_width:
                current_width += step
            if current_width >= self.expand_width:
                self.is_expanding = False
                self.timer.stop()
        elif self.is_shrinking:
            if current_width > self.shrink_width:
                current_width -= step
            if current_width <= self.shrink_width:
                self.is_shrinking = False
                self.timer.stop()

        # Formun ortasını ekranın ortasına göre hesapla
        screen = QtWidgets.QApplication.primaryScreen().availableGeometry()
        x_center = (screen.width() - current_width) // 2
        y_center = (screen.height() - self.fixed_height) // 2
        self.setGeometry(x_center, y_center, current_width, self.fixed_height)
        QtWidgets.QApplication.processEvents()

    def open_excel(self, file_path):
        self.excel = win32com.client.Dispatch("Excel.Application")

        # Excel dosyasını aç
        workbook = self.excel.Workbooks.Open(file_path)
        self.excel.Visible = False  # Excel uygulamasını görünmez yap
        # Excel dosyasındaki makroyu çalıştır
        self.excel.Application.Run("ShowGirisForm")
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.Visible = True  # Excel uygulamasını görünmez yap

    def expand_window(self):
        self.expand_width = 1550  # Genişleme genişliğini ayarla
        self.is_expanding = True
        self.is_shrinking = False
        self.timer.start(10)

    def shrink_window(self):
        self.shrink_width = 1150  # Küçültme genişliğini ayarla
        self.is_expanding = False
        self.is_shrinking = True
        self.timer.start(10)

        sys.exit(app.exec_())

        print(f"Hata yakalandı: {e}")
        # Hata ile başa çıkma veya devam etme işlemleri buraya gelecek

    main()
