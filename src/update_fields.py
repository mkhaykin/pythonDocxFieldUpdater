import logging
import shutil
import subprocess
import sys
import tempfile
from collections.abc import Callable
from pathlib import Path

logger = logging.getLogger(__name__)


class DocxUpdateFieldException(Exception):
    def __init__(
        self,
        message: str = "Ошибка при обновлении полей в документе Word.",
    ) -> None:
        super().__init__(message)


def _file_check(file_path: str) -> Path:
    # Преобразуем путь в объект Path и проверяем его существование
    checked_file_path = Path(file_path).resolve()
    if not checked_file_path.exists() or not checked_file_path.suffix == ".docx":
        raise DocxUpdateFieldException(
            "Указанный файл не существует или имеет неверное расширение",
        )
    return checked_file_path


def _make_backup(file_path: Path) -> Path:
    backup_path = file_path.with_stem(file_path.stem + "_backup")
    shutil.copy2(file_path, backup_path)
    logger.info("Резервная копия сделана.")
    return backup_path


def _win_find_word_app() -> str:
    import winreg

    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
        ) as key:
            word_path, _ = winreg.QueryValueEx(key, "")
            return word_path
    except FileNotFoundError:
        pass

    return "winword"


def _win(file_path: str) -> None:
    import comtypes.client

    cleared_file_path = _file_check(file_path)
    _make_backup(cleared_file_path)

    # TODO добавить проверку на то, что файл уже открыт
    
    word = comtypes.client.CreateObject("Word.Application")
    doc = word.Documents.Open(str(cleared_file_path).replace("/", "\\"))  # TODO smell
    word.ActiveWindow.View.ShowFieldCodes = False
    doc.Repaginate()
    doc.Fields.Update()
    doc.Save()
    word.Quit()
    #
    """
    try:
        command = [
            _win_find_word(),
            file_path,
            "/mFilePrintDefault",
            "/mFileSave",
            "/mFileExit",
            "/q",
        ]
        subprocess.run(command, check=True, shell=False)  # noqa
        logger.info("Поля обновлены успешно.")

    except subprocess.CalledProcessError as e:
        logger.error(f"Ошибка при обновлении полей: {e}")
    """


def _linux(file_path: str) -> None:
    cleared_file_path = _file_check(file_path)
    _make_backup(cleared_file_path)
    temp_path = tempfile.gettempdir()

    libre_path = shutil.which("libreoffice")
    if not libre_path:
        raise DocxUpdateFieldException("Не удалось найти LibreOffice в системе")

    # Используем LibreOffice для обновления полей
    # TODO пересчитывает только страницы, но не арифметические формулы
    try:
        command = [
            str(libre_path),
            "--headless",
            "--invisible",
            "--convert-to",
            "docx",
            '--infilter="MS Word 2010 XML"',
            str(cleared_file_path),
            "--outdir",
            temp_path,
        ]
        subprocess.run(command, check=True, shell=False)  # noqa
        # Возвращаем файл из временного каталога
        shutil.copy2(Path(temp_path, cleared_file_path.name), file_path)
        logger.info("Поля обновлены успешно.")

    except subprocess.CalledProcessError as e:
        logger.error(f"Ошибка при обновлении полей: {e}")
        raise DocxUpdateFieldException()


def _mac(file_path: str) -> None:  # noqa
    # буржуины пока подождут ...
    raise NotImplementedError


def _get_updater() -> Callable:
    if sys.platform == "win32":
        return _win
    elif sys.platform == "linux":
        return _linux
    elif sys.platform == "darwin":
        return _mac

    raise NotImplementedError


def update_fields(
    filename: str,
) -> str:
    return _get_updater()(filename)
