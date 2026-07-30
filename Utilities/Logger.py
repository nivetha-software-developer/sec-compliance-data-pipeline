import inspect
import logging
import os
from typing import Literal, Optional


class Logger:
    ERROR = 'error'
    SUCCESS = 'success'
    WARNING = 'warning'
    INFO = 'info'

    LogLevel = Literal["error", "success", "warning", "info"]

    def __init__(self, automation_name: str, root_folder: str, todayDate):
        log_foler_path = os.path.join(root_folder, 'logs')
        os.makedirs(log_foler_path, exist_ok=True)

        # date folder (logs/20260309)
        log_date_folder = os.path.join(log_foler_path, todayDate)

        os.makedirs(log_date_folder, exist_ok=True)

        self.log_file_name = os.path.join(log_date_folder, f'{automation_name}_{todayDate}.log')
        logging.basicConfig(filename=self.log_file_name, level=logging.INFO, filemode="a",
                                 format='%(asctime)s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        self.log(f'Created logfile at {self.log_file_name}')

    _COLORS = {
        'error': 31,
        'success': 32,
        'warning': 33,
        'info': 36
    }

    _SYMBOLS = {
        'error': '❌',
        'success': '✅',
        'warning': '⚠️',
        'info': 'ℹ️'
    }

    @staticmethod
    def log_to_file(message):
        frame = inspect.currentframe()
        previous_frame = frame.f_back
        previous_frame = previous_frame.f_back

        line_no = previous_frame.f_lineno
        message = f"[Line {line_no}]: {message}"
        logging.info(message)

    def log(self, message: str, level: Optional[LogLevel] = None):
        self.log_to_file(message)
        if level is None:
            print(message)
            return
        try:
            color = Logger._COLORS[level]
            symbol = Logger._SYMBOLS[level]
            print(f"\033[1;{color}m{symbol} {message}\033[0m")
        except KeyError:
            print(f"Unknown level: {level}")


if __name__ == "__main__":
    logger_man = Logger('Test', r'C:\Users\vignesh.r\Desktop\NSE_Derivatives')
    logger, logfilename = logger_man.log, logger_man.log_file_name
    logger(f'Logging at {logfilename}')
    logger('Success', logger_man.SUCCESS)
    logger('Info', logger_man.INFO)
    logger('Warning', logger_man.WARNING)
    logger('Error', logger_man.ERROR)
