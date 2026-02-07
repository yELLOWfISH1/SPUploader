"""
Splash screen for SharePoint Upload Tool
"""
from pathlib import Path

from PySide6.QtCore import Qt, QTimer, QRect, Signal
from PySide6.QtGui import QPixmap, QFont, QColor, QPainter
from PySide6.QtWidgets import QSplashScreen, QApplication


class AnimatedSplashScreen(QSplashScreen):
    """Animated splash screen with progress bar"""

    finished = Signal()

    def __init__(self, pixmap: QPixmap):
        super().__init__(pixmap)
        self.progress = 0
        self.base_pixmap = pixmap.copy()

        self.progress_timer = QTimer()
        self.progress_timer.timeout.connect(self.update_progress)

    def update_progress(self) -> None:
        self.progress += 2
        if self.progress > 100:
            self.progress = 100

        pixmap = self.base_pixmap.copy()
        painter = QPainter(pixmap)

        progress_y = 330
        progress_height = 8
        progress_width = 400
        progress_x = (pixmap.width() - progress_width) // 2

        painter.fillRect(
            progress_x,
            progress_y,
            progress_width,
            progress_height,
            QColor("#2a2a30"),
        )

        fill_width = int((progress_width * self.progress) / 100)
        painter.fillRect(
            progress_x,
            progress_y,
            fill_width,
            progress_height,
            QColor("#c41e3a"),
        )

        painter.setPen(QColor("#555555"))
        painter.drawRect(progress_x, progress_y, progress_width, progress_height)

        percentage_font = QFont("Arial", 9)
        painter.setFont(percentage_font)
        painter.setPen(QColor("#c41e3a"))
        percentage_rect = QRect(progress_x + progress_width + 10, progress_y, 50, progress_height)
        painter.drawText(percentage_rect, Qt.AlignVCenter, f"{self.progress}%")

        painter.end()

        self.setPixmap(pixmap)
        QApplication.processEvents()

        if self.progress >= 100:
            self.progress_timer.stop()
            self.finished.emit()


def create_splash_screen() -> AnimatedSplashScreen:
    """Create and return animated splash screen"""
    pixmap = QPixmap(600, 400)
    pixmap.fill(QColor("#0f1419"))

    painter = QPainter(pixmap)

    logo_path = Path(__file__).resolve().parent.parent / "icons" / "scotiabank_logo_icon_170755.png"
    if logo_path.exists():
        try:
            logo = QPixmap(str(logo_path))
            if not logo.isNull():
                logo = logo.scaledToWidth(150, Qt.SmoothTransformation)
                x = (pixmap.width() - logo.width()) // 2
                y = 50
                painter.drawPixmap(x, y, logo)
        except Exception as exc:
            print(f"Logo load error: {exc}")

    title_font = QFont("Arial", 28, QFont.Bold)
    painter.setFont(title_font)
    painter.setPen(QColor("#ffffff"))

    title_rect = QRect(0, 200, pixmap.width(), 60)
    painter.drawText(title_rect, Qt.AlignCenter, "SharePoint Upload Tool")

    tagline_font = QFont("Arial", 11)
    tagline_font.setItalic(True)
    painter.setFont(tagline_font)
    painter.setPen(QColor("#b0b0b0"))

    tagline_rect = QRect(0, 270, pixmap.width(), 40)
    painter.drawText(tagline_rect, Qt.AlignCenter, "Secure SharePoint Upload Automation")

    loading_font = QFont("Arial", 10)
    painter.setFont(loading_font)
    painter.setPen(QColor("#c41e3a"))
    loading_rect = QRect(0, 350, pixmap.width(), 30)
    painter.drawText(loading_rect, Qt.AlignCenter, "Loading...")

    painter.end()

    splash = AnimatedSplashScreen(pixmap)
    splash.setWindowFlags(splash.windowFlags() | Qt.WindowStaysOnTopHint)
    splash.show()
    QApplication.processEvents()

    splash.progress_timer.start(100)

    return splash
