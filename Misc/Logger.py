
class Logger(object):

    def __init__(self, log_file):
        self._error_count = 0
        self._log_file = log_file
        self._end_line = "---------------------------------------------------------------------------"
        # Delete old contents
        try:
            open(self._log_file, 'w').close()
        except:
            print("Unable to write to log file, verify you have write access to {}".format(self._log_file))

    def log_error(self, message: str):
        try:
            with open(self._log_file, mode="a", encoding="utf-8") as f:
                f.write(message + "\n")
        except:
            pass
        print(message)
        self._error_count += 1

    def log_info(self, message: str):
        try:
            with open(self._log_file, mode="a", encoding="utf-8") as f:
                f.write("~~~" + message + "\n")
        except:
            pass
        print("~~~", message)

    def log_endline(self):
        try:
            with open(self._log_file, mode="a", encoding="utf-8") as f:
                f.write(self._end_line + "\n")
        except:
            pass
        print(self._end_line)

    def print_error_count(self):
        try:
            with open(self._log_file, mode="a", encoding="utf-8") as f:
                f.write("Error Count " + str(self._error_count))
        except:
            pass
        print("Error Count", self._error_count)
