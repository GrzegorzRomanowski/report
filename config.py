import os
from dotenv import load_dotenv
from pathlib import Path

base_dir = Path(__file__).resolve().parent
env_file = base_dir / '.env'
load_dotenv(env_file)


class Config:
    def __init__(self, dd: str, mm: str, yyyy: str, ebawe: str):
        self.TEMPLATE_PATH = Path(os.environ.get('TEMPLATE_PATH'))
        self.EBAWE_REPORT_PATH = Path(os.environ.get('EBAWE_REPORT_PATH').format(DD=dd, MM=mm, YYYY=yyyy, EBAWE=ebawe))
        self.DAILY_PATH = Path(os.environ.get('DAILY_PATH').format(DD=dd, MM=mm, YYYY=yyyy, TEMP=""))
        self.DAILY_TEMP_PATH = Path(os.environ.get('DAILY_PATH').format(DD=dd, MM=mm, YYYY=yyyy, TEMP="_TEMP"))
        self.MONTHLY_PATH = Path(os.environ.get('MONTHLY_PATH').format(MM=mm, YYYY=yyyy, TEMP=""))
        self.MONTHLY_TEMP_PATH = Path(os.environ.get('MONTHLY_PATH').format(MM=mm, YYYY=yyyy, TEMP="_TEMP"))
        self.YEARLY_PATH = Path(os.environ.get('YEARLY_PATH').format(YYYY=yyyy, TEMP=""))
        self.YEARLY_TEMP_PATH = Path(os.environ.get('YEARLY_PATH').format(YYYY=yyyy, TEMP="_TEMP"))
        self.TEMPORARY_FILE = Path(os.environ.get('TEMPORARY_FILE').format(DD=dd, MM=mm, YYYY=yyyy))
