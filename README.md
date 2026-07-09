#**📊 Excel → JSON Converter for BMS PLC Signals**
Это веб-приложение преобразует Excel-файлы (.xlsx/.xls) в структурированный JSON-формат, предназначенный для дальнейшего формирования сигналов PLC в системах Building Management System (BMS) и автоматизации зданий.

Конвертер специально адаптирован под структуру таблиц, содержащих атрибуты тегов PLC: PlcTerminalName, SysObj, SysObjNode, BACNetType, BACNetId, OpcType, VarField, PlcVarDir, Hmi, FsaSchemeName, VarDesc.

##🚀 Возможности
Загрузка Excel-файлов через веб-интерфейс (перетаскивание или выбор файла).

Автоматическое распознавание столбцов по именам (регистр важен).

Преобразование BACNetId в строки (удаление .0).

Обработка булевых значений Hmi (true/false).

Удаление пустых полей FsaSchemeName из выходного JSON.

Форматированный JSON с переносами между объектами (читаемый).

Скачивание результата в виде .json файла.

##🛠️ Требования
Python 3.8+

Установленные библиотеки:

flask

pandas

openpyxl

jinja2

werkzeug

markupsafe

itsdangerous

Все зависимости можно установить одной командой:

bash
pip install -r requirements.txt

##📦 Установка и запуск (в режиме разработки)
Склонируйте репозиторий:

bash
git clone https://github.com/your-username/excel-to-json-bms.git
cd excel-to-json-bms
Создайте виртуальное окружение (рекомендуется):

bash
python -m venv venv
source venv/bin/activate   # Linux/Mac
venv\Scripts\activate      # Windows
Установите зависимости:

bash
pip install -r requirements.txt
Запустите приложение:

bash
python app.py
Откройте браузер по адресу: http://localhost:5000

🏗️ Сборка автономного .exe (для Windows)
Проект включает app.spec для PyInstaller, позволяющий создать один исполняемый файл без необходимости установки Python.

Шаги сборки:
Установите PyInstaller:

bash
pip install pyinstaller
Выполните команду в корневой папке проекта:

bash
pyinstaller app.spec
Готовый файл app.exe появится в папке dist/.

Запуск собранного приложения:
Дважды кликните app.exe.

Автоматически откроется браузер с интерфейсом.

Сервер работает локально на порту 5000.
