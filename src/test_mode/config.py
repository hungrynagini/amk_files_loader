INVALID_CHARS = [[":", ";"], ["\\", "-"], ["|", "-"], ["/", "-"], ["?", "-"], ["*", "-"],
                 [">", "-"], ["<", "-"], ['\t', ""], ['\n', ""], ["\"", "'"]]
SLASH = "/"
PREFIX = ""
import platform
SYSTEM = platform.system()
# print(SYSTEM)
if SYSTEM == "Windows":
    PREFIX = "\\\\?\\"
    SLASH = "\\"
folder = PREFIX
docs_number = 0
docs_done = 0
STOP_EXECUTION = False
METADATA = ['Назва файлу', 'Тип', 'Заголовок', 'Автор', 'Тема', 'Створений', 'Змінений', 'Ким змінено',
            'Програма', 'Виробник пдф', 'Версія пдф', 'Ключові слова', 'Керівник', 'Заклад', 'Архів', 'Папка']
BACKGROUND = '#F0E4E6'
GREEN = '#759C96'
RED = '#C29199'
LIGHTBLUE = '#CFE0F2'
DARKRED = '#711E2C'  #'#78202F' #'#5C0A4D'
