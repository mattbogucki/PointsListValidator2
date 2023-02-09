
class Logger(object):

    def __init__(self, log_file):
        self._error_count = 0
        self._log_file = log_file
        self._end_line = "---------------------------------------------------------------------------"
        # Delete old contents
        open(self._log_file, 'w').close()

    def log_error(self, message: str):
        with open(self._log_file, mode="a", encoding="utf-8") as f:
            f.write(message + "\n")
            print(message)
        self._error_count += 1

    def log_info(self, message: str):
        with open(self._log_file, mode="a", encoding="utf-8") as f:
            f.write("~~~" + message + "\n")
            print("~~~", message)

    def log_endline(self):
        with open(self._log_file, mode="a", encoding="utf-8") as f:
            f.write(self._end_line + "\n")
            print(self._end_line)

    def print_error_count(self):
        with open(self._log_file, mode="a", encoding="utf-8") as f:
            f.write("Error Count " + str(self._error_count))
            print("Error Count", self._error_count)
