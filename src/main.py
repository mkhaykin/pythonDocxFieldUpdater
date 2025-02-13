import logging

from update_fields import update_fields

logger = logging.getLogger(__name__)


def main() -> None:
    update_fields("samples/input.docx")


if __name__ == "__main__":
    main()
