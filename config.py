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


class TestConfig:
    def __init__(self, dd: str, mm: str, yyyy: str, ebawe: str):
        self.TEMPLATE_PATH = Path('test_data/Szablon.xlsx')
        self.EBAWE_REPORT_PATH = Path(f'test_data/E{ebawe}_{dd}.{mm}.{yyyy}.xlsx')
        self.DAILY_PATH = Path(f'test_data/Dzienny_{dd}.{mm}.{yyyy}.xlsx')
        self.DAILY_TEMP_PATH = Path(f'test_data/Dzienny_{dd}.{mm}.{yyyy}_test.xlsx')
        self.MONTHLY_PATH = Path(f'test_data/Miesięczny_{mm}.{yyyy}.xlsx')
        self.MONTHLY_TEMP_PATH = Path(f'test_data/Miesięczny_{mm}.{yyyy}_test.xlsx')
        self.YEARLY_PATH = Path(f'test_data/Roczny_{yyyy}.xlsx')
        self.YEARLY_TEMP_PATH = Path(f'test_data/Roczny_{yyyy}_test.xlsx')
        self.TEMPORARY_FILE = Path('test_data/tymczasowy.xlsx')


config_env = {
    'production': Config,
    'testing': TestConfig
}
