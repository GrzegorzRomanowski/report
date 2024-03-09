import os
from dotenv import load_dotenv
from pathlib import Path

base_dir = Path(__file__).resolve().parent
env_file = base_dir / '.env'
load_dotenv(env_file)


class Config:
    def __init__(self, dd, mm, yyyy, ebawe):
        self.TEMPLATE_PATH = os.environ.get('TEMPLATE_PATH')
        self.EBAWE_REPORT_PATH = os.environ.get('EBAWE_REPORT_PATH').format(DD=dd, MM=mm, YYYY=yyyy, EBAWE=ebawe)
        self.DAILY_PATH = os.environ.get('DAILY_PATH').format(DD=dd, MM=mm, YYYY=yyyy)

