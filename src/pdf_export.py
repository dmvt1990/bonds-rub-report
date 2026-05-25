"""
Конвертация .pptx в .pdf через LibreOffice headless.

На сервере предполагается установленный пакет libreoffice. Проверить:

    which libreoffice
    libreoffice --version

Если LibreOffice не установлен — конвертация будет выбрасывать понятную ошибку,
а сборка pptx всё равно продолжит работать (pdf просто не создастся).
"""
import shutil
import subprocess
from pathlib import Path


class LibreOfficeNotFoundError(RuntimeError):
    """LibreOffice не установлен или не найден в PATH."""


def find_libreoffice() -> str:
    """
    Возвращает путь к исполняемому файлу libreoffice.
    Пробует несколько вариантов имени — разные дистрибутивы ставят по-разному.
    """
    for name in ("libreoffice", "soffice"):
        path = shutil.which(name)
        if path:
            return path
    raise LibreOfficeNotFoundError(
        "LibreOffice не найден в PATH. "
        "Установите: apt install -y libreoffice --no-install-recommends"
    )


def convert_pptx_to_pdf(pptx_path: Path, pdf_output_path: Path, timeout: int = 180) -> Path:
    """
    Конвертирует .pptx в .pdf.

    Под капотом вызывает libreoffice --headless --convert-to pdf.
    LibreOffice умеет писать PDF только в директорию, рядом с
    исходным файлом по имени. Поэтому мы:
      1. Конвертируем в временную папку рядом с pptx
      2. Перемещаем получившийся файл в целевой путь (с нужным именем)

    Args:
        pptx_path: путь к .pptx (должен существовать)
        pdf_output_path: полный путь к будущему .pdf
                         (имя и папка могут быть любыми)
        timeout: максимальное время работы LibreOffice в секундах

    Returns:
        Путь к созданному PDF.

    Raises:
        LibreOfficeNotFoundError: если soffice/libreoffice не установлен
        FileNotFoundError: если pptx_path не существует
        subprocess.CalledProcessError: если LibreOffice вернул ошибку
    """
    pptx_path = Path(pptx_path)
    pdf_output_path = Path(pdf_output_path)

    if not pptx_path.exists():
        raise FileNotFoundError(f"Нет файла pptx: {pptx_path}")

    libreoffice = find_libreoffice()

    # LibreOffice положит pdf в эту директорию с именем <pptx basename>.pdf
    tmp_dir = pptx_path.parent
    expected_pdf = tmp_dir / f"{pptx_path.stem}.pdf"

    # Удаляем старую временную копию, если вдруг осталась от прошлого запуска
    if expected_pdf.exists():
        expected_pdf.unlink()

    cmd = [
        libreoffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(tmp_dir),
        str(pptx_path),
    ]

    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=timeout,
        check=False,
    )

    if result.returncode != 0 or not expected_pdf.exists():
        raise subprocess.CalledProcessError(
            result.returncode, cmd,
            output=result.stdout,
            stderr=result.stderr,
        )

    # Переносим в целевое место с нужным именем.
    # parents=True — создаст /opt/presentation_bot/out/ если её нет.
    pdf_output_path.parent.mkdir(parents=True, exist_ok=True)

    # shutil.move умеет перемещать между разными файловыми системами,
    # а также перезаписывает целевой файл если он есть (нам это и надо).
    if pdf_output_path.exists():
        pdf_output_path.unlink()
    shutil.move(str(expected_pdf), str(pdf_output_path))

    return pdf_output_path
