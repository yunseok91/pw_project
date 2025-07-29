from PyQt5.QtCore import QObject, pyqtSignal
from PyQt5.QtWidgets import QApplication

class Logger(QObject):
    # 시그널 추가
    log_signal = pyqtSignal(str, str)  # 메시지, 타입

    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def log(self, message, type='normal'):
        # 메시지와 스타일 정보를 함께 시그널로 전달
        styles = {
            'normal': 'color: #ffffff',
            'success': 'color: #2ecc71',
            'error': 'color: #e74c3c',
            'warning': 'color: #f39c12',
            'info': 'color: #3498db',
            'separator': 'color: #ffffff'
        }

        style = styles.get(type, styles['normal'])
        base_style = f'font-family: Consolas; font-size: 9pt; {style}'

        if type == 'separator':
            message = "─" * 100
        else:
            emojis = {
                'success': '[Pass] ',
                'error': '[Error] ',
                'warning': '[Warning] ',
                'info': '[INFO] ',
                'normal': '[log] '
            }
            message = f"{emojis.get(type, '')}{message}"

        # 포맷된 HTML 메시지를 시그널로 전달
        formatted_message = f'<span style="{base_style}">{message}</span><br>'
        self.log_signal.emit(formatted_message, type)