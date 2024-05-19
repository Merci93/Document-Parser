from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    file_location: str = "../files_to_parse"
    parsed_data_directory: str = "../parsed_files"
    extracted_images: str = "images"
    extracted_images_file_name: str = "images.csv"
    extracted_table_of_content: str = "table_of_content"
    extracted_tables: str = "tables"
    extracted_texts: str = "texts"
    parse_report: str = "parse_report.csv"


def init_settings() -> None:
    global settings
    settings = Settings()


settings = None
if not settings:
    init_settings()
